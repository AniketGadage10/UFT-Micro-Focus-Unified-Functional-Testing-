'1. write a program to login SauceDemo.com and add any item in shopping cart and remove then logout. Use visual relationship for adding or of add to cart button.
Option Explicit

Dim BRW_Obj

Set BRW_Obj=Browser("BRW_SAUCEDEMO").Page("BPG_SAUCEDEMO")
SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://www.saucedemo.com/"

	BRW_Obj.WebEdit("user-name").Set "standard_user"
	BRW_Obj.WebEdit("user_password").Set "secret_sauce"
	BRW_Obj.WebButton("Login").Click
	
	
Browser("BRW_SAUCEDEMO").Sync

Browser("BRW_SAUCEDEMO").InsightObject("BAG_Add_To_Cart").Click
Browser("BRW_SAUCEDEMO").InsightObject("t-shirt_Add_TO_Card").Click

Browser("BRW_SAUCEDEMO").Page("Porduct_Swag Labs").WebElement("Shopping_Card").Click

Browser("BRW_SAUCEDEMO").InsightObject("Remove_Bag").Click

Browser("BRW_SAUCEDEMO").Page("Your_Card").WebButton("User_Open Menu").Click

Browser("BRW_SAUCEDEMO").Page("Your_Card").Link("User_Logout").Click

