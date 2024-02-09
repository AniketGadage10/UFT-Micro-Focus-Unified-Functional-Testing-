'Prerquisite - "c:\MyFolder\Trainee_Names.xlsx"
'->Sheet1 - available data -> cell(2,1) = "Hrushikesh" , cell(4,1) = "Prajakta" , Cell(5,1) = "Ranajeet" , Cell(3,1)= "Siddharth"
'Write a vb script which will read above excel and create folders at "c:\MyFolder\" location as per Trainee Names
'Expected O/P : after executing .vbs 4 different folders should get created at "c:\MyFolder\" location with folder names Hrushikesh,Prajakta,Ranajeet,Siddharth

Option Explicit

Dim File_Obj

Dim xl_obj,xl_Workbook,xl_worksheet

Dim src,dest,NEW_PATH

Dim irow_count

Dim int_i:int_i=0

dest="c:\MyFolder\AniketGadage"

ForRead=1

src="c:\MyFolder\Trainee_Names.xlsx"

Set xl_Obj=CreateObject("excel.Application")
xl_Obj.visible=True
Set xl_Workbook=xl_obj.Workbooks.open(src)

Set xl_worksheet=xl_Workbook.worksheets(1)

xl_worksheet.cells(2,1).value = "Hrushikesh" 
xl_worksheet.cells(4,1).value = "Prajakta" 
xl_worksheet.cells(5,1).value  = "Ranajeet"  
xl_worksheet.cells(3,1).value = "Siddharth"

irow_count=xl_worksheet.usedrange.ROWS.Count

File_Obj.CreateFolder dest

For int_i=2 to irow_count
	NEW_PATH=dest+"\"+xl_worksheet.cells(int_i,1).value
	File_Obj.CreateFolder NEW_PATH
Next

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing


