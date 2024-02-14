'Question 6 
'Automate below scenario. use function library to store all functions
'Library 1( LaunchApplication , logoutFromApplication , Login ) 
'Library 2( updateContactInfo ,RegisterNewUser )
'Step 1 : open https://parabank.parasoft.com/parabank/index.htm on chrome
'Step 2 : Click on "Register" link
'Step 3 : Fill the required details and click on register
'Step 4 : Click on logout link
'Step 5 : Login with created user and Password
'Step 6 : click on update contact info and click on Update Profiles


LaunchedApplication()

Browser("ParaBank | Welcome | Online").Page("ParaBank | Welcome | Online").Link("Register").Click
		
		'RegisterNewUser DataTable.Value("sFName","Global"),DataTable.Value("sLName","Global").DataTable.Value("sAdd","Global"),DataTable.Value("sCity","Global"),DataTable.Value("sState","Global"),DataTable.Value("sZipCode","Global"),DataTable.Value("sPhone","Global"),DataTable.Value("sSSN","Global"),DataTable.Value("sUserName","Global"),DataTable.Value("sPassword","Global")
		
	if  RegisterNewUser("Aniket","Gadage","Shivaji Chowk","Pune","Maharashtra",	"411057","7741029614","1236544","Aniket10","Aniket@123") then
		print "RegisterNewUser Sucessfull"
		
		Login "Aniket10","Aniket@123"
		wait(1)
		updateContactInfo "Aniket1","Gadage1","Shivaji Chowk1","Pune1","Maharashtra1",	"411057","7741029614"
		logoutFromApplication()
	else
		print "RegisterNewUser Fail"
	End If
	
