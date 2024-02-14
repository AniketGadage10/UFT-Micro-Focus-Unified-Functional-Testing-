'Question 1
'Automate above manual test case with the help of normal recording. 
'Scenario should run without any issue when we rerun recorded script. Use Sync concepts if needed



SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://demo.automationtesting.in/Index.html"

Browser("Index").Page("Index").WebEdit("Email id for Sign Up").Set "aniketgadage1018@gmail.com"
Browser("Index").Page("Index").Image("logo").Click
Browser("Index").Page("Register").WebEdit("First Name").Set "Aniket" @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Index").Page("Register").WebEdit("Last Name").Set "Gadage" @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Index").Page("Register").WebEdit("WebEdit").Set "Pune" @@ script infofile_;_ZIP::ssf5.xml_;_
Browser("Index").Page("Register").WebEdit("WebEdit_2").Set "aniketgadage1018@gmail.com" @@ script infofile_;_ZIP::ssf6.xml_;_
Browser("Index").Page("Register").WebEdit("WebEdit_3").Set "7741029614" @@ script infofile_;_ZIP::ssf7.xml_;_
Browser("Index").Page("Register").WebRadioGroup("radiooptions").Select "Male" @@ script infofile_;_ZIP::ssf8.xml_;_
Browser("Index").Page("Register").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("Index").Page("Register").WebElement("msdd").Click @@ script infofile_;_ZIP::ssf10.xml_;_
Browser("Index").Page("Register").WebElement("English").Click @@ script infofile_;_ZIP::ssf11.xml_;_
Browser("Index").Page("Register").WebList("select").Select "Java" @@ script infofile_;_ZIP::ssf12.xml_;_
Browser("Index").Page("Register").WebList("WebList").Click @@ script infofile_;_ZIP::ssf13.xml_;_
Browser("Index").Page("Register").WebList("select_2").Select "2000" @@ script infofile_;_ZIP::ssf14.xml_;_
Browser("Index").Page("Register").WebList("select_3").Select "March" @@ script infofile_;_ZIP::ssf15.xml_;_
Browser("Index").Page("Register").WebList("select_4").Select "10" @@ script infofile_;_ZIP::ssf16.xml_;_
Browser("Index").Page("Register").WebEdit("WebEdit_4").SetSecure "65cc47047ad58cb522cd192c1de673a5552c94ce5c5a1e13cc138ca6" @@ script infofile_;_ZIP::ssf17.xml_;_
Browser("Index").Page("Register").WebEdit("WebEdit_5").SetSecure "65cc470e1182f35386aae2470396aecb8941ce01aa6313d4fdc63a2c" @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("Index").Page("Register").WebButton("Refresh").Click

Browser("Index").Close
