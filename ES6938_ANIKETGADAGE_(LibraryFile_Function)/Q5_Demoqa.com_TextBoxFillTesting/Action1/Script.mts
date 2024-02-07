'5. Write a Program to open Website https://demoqa.com/ and 
'in Elements on Textbox tab fill details in edit boxes and submit. write function to perform operation and verify text after submit


Browser("DEMOQA").Page("DEMOQA").Sync

Browser("DEMOQA").Page("DEMOQA").WebElement("WebElement").Click

	If Browser("DEMOQA").Page("DEMOQA_2").WebElement("item-0").Exist(10) Then
	
		TextFill DataTable("Full_Name"),DataTable("Email"),DataTable("Current_Address"),DataTable("Permanent_Address") 
		
		If  ValidateTextBox( DataTable("Full_Name"),DataTable("Email"),DataTable("Current_Address"),DataTable("Permanent_Address")   ) Then
			Print "Validation Sucessfull"	
		Else
			Print "Validation Fail"
		End If
	Else
		Print "Error In Loading Text-Box "
	End If
