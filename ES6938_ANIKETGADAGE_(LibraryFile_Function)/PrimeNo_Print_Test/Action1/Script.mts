'Q1 Write a test case print prime no of given range (eg, 20 to 60) Using function having parameter startRange and EndRange




'------------------------------------------------------Load Function Libary File--------------------------------------------------------------------------------------------
'LoadFunctionLibrary "C:\Users\agadage\Desktop\UFT-Micro-Focus-Unified-Functional-Testing-\ES6938_ANIKETGADAGE_(LibraryFile_Function)\LibraryFunction\PrimeNo_Functions.vbs"

'------------------------------------------------------USING HARD CODE  INPUT--------------------------------------------------------------------------------------------
'Dim Start_no,End_No
'
'Start_No=20
'End_No=60
'
'MsgBox GetPrimeNumber_In_Range(Start_No,End_No)
'
'Print GetPrimeNumber_In_Range(Start_No,End_No)

'------------------------------------------------------USING DATA TABLE INPUT--------------------------------------------------------------------------------------------

Dim Start_no,End_No

MsgBox GetPrimeNumber_In_Range(DataTable("Start_No"),DataTable("End_No"))

Print GetPrimeNumber_In_Range(DataTable("Start_No"),DataTable("End_No"))
