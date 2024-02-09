'8. (10 Marks)
'A. (5 Marks)
'Write a program to calculate total number of months difference between below dates
'dtDate1 = #07 Jan 2015#
'dtDate2 = #08 Jun 2016#
'dtDate3 = #10 Apr 2017#
'dtDate4 = #09 Sep 2019#

Option Explicit

Dim dtDate1,dtDate2,dtDate3,dtDate4

Dim Month_Sum:Month_Sum=0

dtDate1 = #07 Jan 2015#
dtDate2 = #08 Jun 2016#
dtDate3 = #10 Apr 2017#
dtDate4 = #09 Sep 2019#

Month_Sum=DateDiff("m",dtDate1,dtDate2)+DateDiff("m",dtDate2,dtDate3)+DateDiff("m",dtDate3,dtDate4)

MsgBox "Month_Sum = "&Month_Sum