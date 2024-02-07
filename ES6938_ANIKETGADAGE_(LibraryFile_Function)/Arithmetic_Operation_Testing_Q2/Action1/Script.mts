' write a program to perform arithmetic operation eg. add, sub, mult, div. create library and write function having parameter action, FirstNum, SecondNum.
Dim IRow_Count
Dim Iterator

IRow_Count=DataTable.GetRowCount

For Iterator = 1 To IRow_Count Step 1
	Print DataTable("FirstNum")&"  "&DataTable("Action")&"  "&DataTable("SecondNum")&"  =  "&Calculator(DataTable("Action"),DataTable("FirstNum"),DataTable("SecondNum"))
	DataTable.SetNextRow
Next
