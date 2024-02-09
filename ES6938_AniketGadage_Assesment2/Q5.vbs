'5. (10 Marks)
'Write a vb script program to accept Car's brand name and car name  from user 
'based on brand and Car, Prize should get populated 
'(e.g TATA - Safari 25L,Nexon 15L,Punch 10L,Tigor 8L, Mahindra - XUV700 19 L,XUV500 15L,XUV300 12 L
' , Honda - Jazz 9L,Amaze 10L,City 12L, Skoda - Kushaq 17 L,Slavia 18L,Octavia 20 L) 


Option Explicit

Dim Car_Brand,Car_Name
Dim Concat_Name

Car_Brand=LCase(InputBox("Enter The Car Brand Name"))
Car_Name=LCase(InputBox("Enter The Car Name"))


Select Case Car_Brand

Case "tata"

	If Car_Name="safari" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 25L"
	ELSEIF Car_Name="nexon" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 15L"
	ELSEIF Car_Name="punch" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 10L"
	ELSEIF Car_Name="tigor" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 8L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

Case "mahindra"
	If Car_Name="xuv700" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 19L"
	ELSEIF Car_Name="xuv500" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 15L"
	ELSEIF Car_Name="xuv300" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 12L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

Case "honda"
	If Car_Name="jazz" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 9L"
	ELSEIF Car_Name="amaze" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 10L"
	ELSEIF Car_Name="city" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 12L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

Case "skoda"
	If Car_Name="kushaq" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 17L"
	ELSEIF Car_Name="slavia" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 18L"
	ELSEIF Car_Name="octavia" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 20L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

End Select