'write test case to login ang logout SauceDemo.com site using functions username standard user and password secret sauce

 Dim  Valid_User,Valid_Password

Valid_User="standard_user"
Valid_Password="secret_sauce"

IF Authauthentication(Valid_User,Valid_Password) Then
		
		Print "User Login Sucessfully"
		
		If Logout Then
		
			Print "User Login Out Sucessfully"
			
		Else
			Print "User Login UnSucessfully"
		End If
Else
	Print "User Login Fail"
End  If

'performance_glitch_user
'error_user
'visual_user
