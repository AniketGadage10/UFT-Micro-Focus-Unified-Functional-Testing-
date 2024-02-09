'9. (10 Marks)
'Write a program to handle below erros for uninterrupted execution. get the error number and description as well for fail log

'Option Explicit
'intVal1  =  take input from user 'eg 2
'intVal2  =  take input from user 'eg 2
'intVal2 = intVal1 - intVal2 
'intVal2 = intVal1/intVal2
'msgbox intVal2



Option Explicit


Dim intVal1,intVal2

ON error Resume NEXT

	intVal1  =  Int(InputBox("ENTER THE NUMBER 1"))
	intVal2  =  Int(InputBox("ENTER THE NUMBER 2"))
	intVal2 = intVal1 - intVal2 
	intVal2 = intVal1/intVal2

	if Err.number<>0 Then
		msgbox "ERROR number = "&Err.number
		msgbox "ERROR description = "&Err.description
	Else
		msgbox "Code Run Sucessfull"
	end if

Err.clear

ON error Goto 0
