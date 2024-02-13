SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://iyku0.csb.app/"

SET obj=Browser("React App").Page("React App").WebTable("WEBTABLE")

print obj.RowCount

print obj.ColumnCount(1)

print obj.GetCellData(1,2)

print obj.GetRowWithCellText("Cheese",3,1)

Set Odesc =Description.Create

Odesc("html tag").Value="Button"

print obj.ChildItemCount(2,4,"WebButton")

obj.ChildItem(2,4,"WebButton",0).Click

Set child_obj=Browser("React App").Page("React App").WebElement("ActionsUpdate Delete").ChildObjects(Odesc)

PRINT child_obj.count

For Iterator = 0 To child_obj.count-1 Step 1
	
	print child_obj(Iterator).GetRoproperty("innertext")
Next

