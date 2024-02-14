'Question 9
'Step 1 : launch chrome and go to https://demo.automationtesting.in/DynamicData.html
'Step 2 : Click on get dynamic data button
'Step 3 : get First Name and Last name from application and store it in data table


SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://demo.automationtesting.in/DynamicData.html"

Browser("File input - Multi select").Page("File input - Multi select").WebButton("Get Dynamic Data").Click

Set odesc=Description.Create

odesc("html id").Value="loading"


Set Child =Browser("File input - Multi select").Page("File input - Multi select").WebElement("Clicking on 'Get Dynamic").ChildObjects(odesc)

print Child.Count

print Child(0).GetRoProperty("innertext")

