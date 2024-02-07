
SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://demoqa.com/webtables"


'Check data availability in webtable
Set objDemoQAWebtable = Browser("brw_DemoQA").Page("wpg_DemoQA").WebTable("wbtbl_DemoQA")

if objDemoQAWebtable.Exist(3) Then
	Print "Webtable Present"
Else
	Print "Fail : Webtable not Present"
End  If

Print objDemoQAWebtable.RowCount
Print objDemoQAWebtable.ColumnCount(2)
Print objDemoQAWebtable.GetCellData(1,1)
