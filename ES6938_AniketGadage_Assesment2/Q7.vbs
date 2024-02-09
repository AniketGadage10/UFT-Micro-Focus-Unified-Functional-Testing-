'7. (10 Marks)
'Write a program to replace 3rd & 4th Expleo/expleo word with "Expleo India" . Make sure while showing output 
'msgbox should show entire string with replaced words, and not the half string from where you started replacing
'strOrgDetails = "Expleo office Pune. expleo head count is 2K. I work in Expleo. expleo is very nice organisation"


Option Explicit

Dim strOrgDetails,TEMP_LEFT,TEMP_RIGHT,ch
Dim Search_Count,int_i
strOrgDetails = "Expleo office Pune. expleo head count is 2K. I work in Expleo. expleo is very nice organisation"
Search_Count=0
TEMP_LEFT=""
TEMP_RIGHT=""

For int_i=1 to Len(strOrgDetails)
	ch=Mid(strOrgDetails,int_i,6)
	
	IF ch="Expleo" Or Ch="expleo" Then
		Search_Count=Search_Count+1
		IF Search_Count=3 OR Search_Count=4 Then
			TEMP_LEFT=Left(strOrgDetails,int_i-1)
			TEMP_RIGHT=Mid(strOrgDetails,int_i+6,Len(strOrgDetails)-Len(TEMP_LEFT)-6)
			strOrgDetails=TEMP_LEFT+"Expleo India"+TEMP_RIGHT
		end IF
	end if	
Next

MsgBOX strOrgDetails