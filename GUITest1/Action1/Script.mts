'LoadFunctionLibrary "C:\MyFolder\Function\Library1.qfl" @@ hightlight id_;_593234_;_script infofile_;_ZIP::ssf11.xml_;_


LOGIN()


'Desktop.RunAnalog "Track1"

'Window("Micro Focus MyFlight Sample").Click 195,228 @@ hightlight id_;_4065650_;_script infofile_;_ZIP::ssf12.xml_;_
'Window("Micro Focus MyFlight Sample").Type "JOHN" @@ hightlight id_;_4065650_;_script infofile_;_ZIP::ssf13.xml_;_
'Window("Micro Focus MyFlight Sample").Click 266,292 @@ hightlight id_;_4065650_;_script infofile_;_ZIP::ssf14.xml_;_
'Window("Micro Focus MyFlight Sample").Type "HP" @@ hightlight id_;_4065650_;_script infofile_;_ZIP::ssf15.xml_;_
'Window("Micro Focus MyFlight Sample").Click 185,342 @@ hightlight id_;_4065650_;_script infofile_;_ZIP::ssf16.xml_;_
'Window("Micro Focus MyFlight Sample_2").Close

'print WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").GetTOProperties("devname")
'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").SetTOProperty "devname","anu"
'print WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").GetTOProperties("devname")
'wait(1)

'RunAction "Action2",oneiteration,2,3

'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("devname:=agentName").Set "jOHN"



