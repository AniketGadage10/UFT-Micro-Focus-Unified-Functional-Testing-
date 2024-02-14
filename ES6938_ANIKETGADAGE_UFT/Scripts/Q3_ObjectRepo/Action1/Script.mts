'Question 3
'Automate above manual test case by adding objects in shared object repository. Follow all standards for adding objects in OR
'follow all coding standards while writting code

SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://demo.automationtesting.in/Index.html"

Browser("Index").Page("Index").Sync

Browser("Index").Page("Index").WebEdit("Email id for Sign Up").Set "aniketgadage1018@gmail.com"
Browser("Index").Page("Index").Image("logo").Click

Browser("Index").Page("Register").Sync

Browser("Index").Page("Register").WebEdit("First Name").Set "Aniket"
Browser("Index").Page("Register").WebEdit("Last Name").Set "Gadage"
Browser("Index").Page("Register").WebEdit("WebEdit").Set "Pune"
Browser("Index").Page("Register").WebEdit("WebEdit_2").Set "aniketgadage1018@gmail.com"
Browser("Index").Page("Register").WebEdit("WebEdit_3").Set "7741029614"
Browser("Index").Page("Register").WebRadioGroup("radiooptions").Select "Male"
Browser("Index").Page("Register").WebCheckBox("WebCheckBox").Set "ON"
Browser("Index").Page("Register").WebElement("msdd").Click
Browser("Index").Page("Register").WebElement("English").Click
Browser("Index").Page("Register").WebList("select").Select "Java"
Browser("Index").Page("Register").WebList("WebList").Click
Browser("Index").Page("Register").WebList("select_2").Select "2000"
Browser("Index").Page("Register").WebList("select_3").Select "March"
Browser("Index").Page("Register").WebList("select_4").Select "10"
Browser("Index").Page("Register").WebEdit("WebEdit_4").SetSecure "65cc47047ad58cb522cd192c1de673a5552c94ce5c5a1e13cc138ca6"
Browser("Index").Page("Register").WebEdit("WebEdit_5").SetSecure "65cc470e1182f35386aae2470396aecb8941ce01aa6313d4fdc63a2c"

Browser("Index").Page("Register").WebButton("Refresh").Click

Browser("Index").Close

