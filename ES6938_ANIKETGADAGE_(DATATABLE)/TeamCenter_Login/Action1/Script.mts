'TeamCenter Login

'step 1:Clear the InputBoxes
JavaWindow("Teamcenter Login").JavaButton("Clear").Click

'step 2:Enter The User Id

JavaWindow("Teamcenter Login").JavaEdit("Jedit_UserID").Set DataTable.Value("Username","Action1")

'step 3:Enter The Password

JavaWindow("Teamcenter Login").JavaEdit("Jedit_Password").Set DataTable.Value("Password","Action1")

'step 4:Click Login Button
JavaWindow("Teamcenter Login").JavaButton("Login").Click

Wait(10)
'Step 5: Create New Folder Under Home

JavaWindow("My Teamcenter - Teamcenter").JavaMenu("File").JavaMenu("New").JavaMenu("Folder...").Select

'Step 5:Click On Next Button
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("Folder_Selection_Next_Button").JavaButton("Next >").Click

'Step 5:Insert Folder Name
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("Folder_Selection_Next_Button").JavaEdit("Folder_Name_Input").Set DataTable.Value("Folder_Name","Action1")
Wait(2)
'Step 6: Click On Finish
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("Folder_Selection_Next_Button").JavaButton("Finish").Click

'Step 7: Close The Tab
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("Folder_Selection_Next_Button").JavaButton("Close").Click

'Step 8 : 
JavaWindow("My Teamcenter - Teamcenter").InsightObject("InsightObject").Click

'Step 9:Add Item Under  Folder

JavaWindow("My Teamcenter - Teamcenter").JavaMenu("File").JavaMenu("New").JavaMenu("Item...").Select

'Step 10 : ITEM Click On Next Button
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("File_Selection_Next_Button").JavaButton("Next >").Click

'Step 11: Insert File Name
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("File_Selection_Next_Button").JavaEdit("File_Name_Input").Set  DataTable.Value("File_Name","Action1")

'Step 12: Click On Finish
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("File_Selection_Next_Button").JavaButton("Finish").Click

'Step 13: Close The Tab
JavaWindow("My Teamcenter - Teamcenter").JavaWindow("File_Selection_Next_Button").JavaButton("Close").Click

'Step 14 : Exit Application
JavaWindow("My Teamcenter - Teamcenter").JavaMenu("File").JavaMenu("Exit").Select
