'6. (10 Marks)
'Write a vb script to create a new excel and store below table.
'use minimum for loops
'1	2	3	4	5
'2	4	6	8	10
'3	6	9	12	15
'4	8	12	16	20
'5	10	15	20	25
'6	12	18	24	30
'7	14	21	28	35
'8	16	24	32	40
'9	18	27	36	45
'10	20	30	40	50




Option Explicit

Dim File_Obj

Dim xl_obj,xl_Workbook,xl_worksheet

Dim src,dest,NEW_PATH

Dim irow_count,icol_count

Dim int_i:int_i=0

dest="c:\MyFolder\AniketGadage"

ForRead=1
irow_count=10
icol_count=5

src="c:\MyFolder\Trainee_Names.xlsx"

Set xl_Obj=CreateObject("excel.Application")

xl_Obj.visible=True

Set xl_Workbook=xl_obj.Workbooks.open(src)

Set xl_worksheet=xl_Workbook.worksheets(1)

	for int_i=1 to icol_count
			for int_j=1 to irow_count
				xl_worksheet.cells(irow_count,icol_count).value = icol_count*irow_count
			Next
	Next

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing






	
Dim