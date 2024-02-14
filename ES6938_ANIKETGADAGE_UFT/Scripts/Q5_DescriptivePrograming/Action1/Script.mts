'Question 5
'Automate above manual test case with Descriptive programming. Follow standards to declare objects and memory clear
'follow all coding standards while writting code


SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://demo.automationtesting.in/Index.html"

wait(2)

Browser("Index").Page("Index").WebEdit("placeholder:=Email id for Sign Up").Set "aniketgadage1018@gmail.com"

Browser("Index").Page("Index").Image("html id:=enterimg").Click

Browser("Index").Page("Register").WebEdit("placeholder:=First Name").Set "Aniket"

Browser("Index").Page("Register").WebEdit("placeholder:=Last Name").Set "Gadage"

Browser("Index").Page("Register").WebEdit("html tag:=TEXTAREA").Set "Pune"

Browser("Index").Page("Register").WebEdit("type:=email").Set "aniketgadage1018@gmail.com"

Browser("Index").Page("Register").WebEdit("type:=tel").Set "7741029614"

Browser("Index").Page("Register").WebRadioGroup("name:=radiooptions").Select "Male"

Browser("Index").Page("Register").WebCheckBox("html id:=checkbox1").Click @@ script infofile_;_ZIP::ssf1.xml_;_

Browser("Index").Page("Register").WebElement("html id:=msdd").Click

Browser("Index").Page("Register").WebElement("innerhtml:=English").Click

Browser("Index").Page("Register").WebList("value:=Select Skills").Select "Java"

Browser("Index").Page("Register").WebList("class:=select2-selection select2-selection--single").Click @@ script infofile_;_ZIP::ssf8.xml_;_

Browser("Index").Page("Register").WebEdit("type:=search").Set "india"

Browser("Index").Page("Register").WebButton("name:=Refresh").Click
 @@ script infofile_;_ZIP::ssf9.xml_;_

Browser("Index").Close


