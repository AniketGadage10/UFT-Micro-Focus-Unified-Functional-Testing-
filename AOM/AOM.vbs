'Creating UFT One Application object

	Set objUFT = CreateObject("QuickTest.Application")

	If objUFT.Launched = False Then

		'Launching UFT One and applying basic settings

		objUFT.Launch

		objUFT.Visible = True
		objUFT.Open "C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\Q3_TestDestails_LogFile_Creation",True,False
		Set uftLibraries = objUFT.Test.Settings.Resources.Libraries
		'If the library file "libraary.vbs" is not assiciates with the Test then associate it
			If uftLibraries.Find("C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\LibraryFunction\Arithmetic_Calculator_Function.vbs") = -1 Then
				uftLibraries.Add "C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\LibraryFunction\Arithmetic_Calculator_Function.vbs", 1
			End If
			
	    Set uftRepositories = objUFT.Test.Actions("Action1").ObjectRepositories
		' Add  Object repositry "Reposit.tsr" if it’s not already associated wit action "SignIn"
		If uftRepositories.Find("C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(ObjectRepo)\SharedRepo\Shared_Repo1.tsr") = -1 Then
			uftRepositories.Add "C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(ObjectRepo)\SharedRepo\Shared_Repo1.tsr", 1
		End If		
		'Save the test
		objUFT.Test.Save
		
		objUFT.Test.Run
		msgbox "OK"
		objUFT.Quit
    End If		