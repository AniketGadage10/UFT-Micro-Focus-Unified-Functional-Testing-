
Option Explicit
SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://www.onlinecalculator.com/"
Dim irow	

For irow = 1 To 	DataTable.GetRowCount Step 1

DataTable.SetCurrentRow(irow)

Select Case  DataTable.Value("Operation","Global")

	Case "+"
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").SetTOProperty "name", DataTable.Value("NUM1","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").SetTOProperty "name", DataTable.Value("Operation","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").SetTOProperty "name", DataTable.Value("NUM2","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").Click	
			
			
	Case "-"
	
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").SetTOProperty "name", DataTable.Value("NUM1","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").SetTOProperty "name", DataTable.Value("Operation","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").SetTOProperty "name", DataTable.Value("NUM2","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").Click	
			
	
	Case "X"
	
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").SetTOProperty "name", DataTable.Value("NUM1","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").SetTOProperty "name", DataTable.Value("Operation","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").SetTOProperty "name", DataTable.Value("NUM2","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").Click	
			

	
	Case "/"
	
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").SetTOProperty "name", DataTable.Value("NUM1","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_1").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").SetTOProperty "name", DataTable.Value("Operation","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Operation").Click
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").SetTOProperty "name", DataTable.Value("NUM2","Global")
			Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Var_2").Click	
			
	
End Select

Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("EqualsTo").Click
print  DataTable.Value("NUM1","Global")&"   "& DataTable.Value("Operation","Global")&"  "&DataTable.Value("NUM2","Global")&"  =   "& Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebElement("Calculator_OutputBox").GetROProperty("innertext")
Browser("Online Calculator - OnlineCalc").Page("Online Calculator - OnlineCalc").WebButton("Clear").Click


Next
