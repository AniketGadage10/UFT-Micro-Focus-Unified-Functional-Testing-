
Dim objUFT
Dim bReturn

	'Terminating unwanted process

	bReturn = Fn_BatchRunner_KillProcess("javaw.exe")

	bReturn = Fn_BatchRunner_KillProcess("UFT Onero.exe")

	bReturn = Fn_BatchRunner_KillProcess("QTAutomationAgent.exe")	

	If bReturn = True Then

		'Creating UFT One Application object

		Set objUFT = CreateObject("UFT.Application")

		If objUFT.Launched = False Then

			'Launching UFT One and applying basic settings

			objUFT.Launch

			objUFT.Visible = True

			objUFT.Test.Settings.Run.ObjectSyncTimeOut = 60			

			objUFT.Test.Settings.Java.EditSet.KeyPressed = True

			objUFT.Test.Settings.Java.EditSet.KeyTyped = True

			objUFT.Test.Settings.Java.EditSet.KeyReleased = True

			objUFT.Test.Settings.Java.EditSet.InputMethod = True

			objUFT.Test.Settings.Java.EditSet.FocusLost = False

			objUFT.Test.Settings.Java.EditSet.JavaAPI = True

			objUFT.Test.Settings.Java.EditSet.JavaAPICharactersToExclude = "\t\n\r"

			objUFT.Test.Settings.Java.DeviceReplay.Click = True

			objUFT.Test.Settings.Java.DeviceReplay.DoubleClick = True

			objUFT.Test.Settings.Java.DeviceReplay.Type = True

			objUFT.Test.Settings.Java.DeviceReplay.DragDrop = True					

			objUFT.WindowState = "Maximized"

			objUFT.ActivateView "ExpertView"

			objUFT.Open sScriptName,True,False

			objUFT.Test.Run

			objUFT.Quit

		
		'Releasing UFT One Application object	

		Set objUFT = Nothing

		


