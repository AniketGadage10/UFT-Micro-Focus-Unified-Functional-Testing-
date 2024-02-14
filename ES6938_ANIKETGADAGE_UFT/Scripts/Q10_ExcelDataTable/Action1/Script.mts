'Question 10
'Import Excel which has below data in Sheet1. After importing perform multiplication of first and Second Coulmn
'and store the result in Action1 Datatable sheet in Result column . Export Result in excel 
'FirstIntVal SecIntVal
'	5			6
'	3			8
'	9			2

Option Explicit

Dim iRowCount,Int_i,Result

DataTable.ImportSheet "C:\Users\agadage\Desktop\ES6938_ANIKETGADAGE_UFT\TestData\DatatableOperation.xlsx",1,"Global"

iRowCount=DataTable.GlobalSheet.GetRowCount

For  Int_i = 1 To iRowCount Step 1
	DataTable.SetCurrentRow(Int_i)
	Msgbox DataTable.Value("FirstIntVal","Global")
	Result=Int(DataTable.Value("FirstIntVal","Global"))*Int(DataTable.Value("SecIntVal","Global"))
	DataTable.Value("Result","Global")=Result
Next

DataTable.ExportSheet "C:\Users\agadage\Desktop\ES6938_ANIKETGADAGE_UFT\TestData\Result.xlsx","Global"

DataTable.ImportSheet "C:\Users\agadage\Desktop\ES6938_ANIKETGADAGE_UFT\TestData\Result.xlsx",1,"Action1"

Wait(2)
