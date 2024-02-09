'*********************************************************	Function List		***********************************************************************
'1. Fn_SendEmail(strSendTo, strSendCC, strSubject, strTextBody, strAttachmentPath)
'*********************************************************	Function List		***********************************************************************

Function Fn_SendEmail(strSendTo, strSendCC, strSubject, strTextBody, strAttachmentPath)
	Dim Iterator, attachmentsArr
	Set objMessage = CreateObject("CDO.Message")
	
	'==This section provides the configuration information for the remote SMTP server.
	'==Normally you will only change the server name or IP.
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "pnsmtp.ugs.com"
	
	'Server port (typically 25)
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	objMessage.Configuration.Fields.Update
	objMessage.Subject = strSubject
	objMessage.From = "tcautomation@siemens.com"
	objMessage.To = strSendTo
	objMessage.CC = strSendCC
	'objMessage.TextBody = strTextBody
    objMessage.HTMLBody = strTextBody
    
    attachmentsArr = split(strAttachmentPath,";")
    For Iterator = 0 To UBound(attachmentsArr) Step 1
    	objMessage.AddAttachment attachmentsArr(Iterator)
    Next
	
	objMessage.Send
	
	Set objMessage = Nothing

	Fn_SendEmail = True

End Function

'*********************************************************	Function List		***********************************************************************
'2. Fn_EmailBodyTxt(strExcelPath)
'*********************************************************	Function List		***********************************************************************
Function Fn_EmailBodyTxt(strExcelPath)

	Dim bReturn
	Dim sBatchResult
	Dim sBatchDate
	Dim sBatchDuration
	Dim sBatchLogFolder
	Dim sBatchTotalTestCases
	Dim sBatchPassTestCases
	Dim sBatchFailTestCases
	Dim sMailBody

	' Batch result
	bReturn = Fn_ExcelGetResultDetail(strExcelPath,"FailExist", 1)
	if bReturn = "True" Then
		sBatchResult = "FAIL" 
	else
		sBatchResult = "PASS"
	end if	

	' Batch date
	sBatchDate = Fn_BatchDate(strExcelPath)

	' Batch duration
	sBatchDuration = Fn_BatchDuration(strExcelPath)

	' Batch log folder
	sBatchLogFolder = Environment.Value("LocalHostName") & "\" &Fn_GetBatchName(Environment("BatchFldName"))

	' Total testcases in batch
	sBatchTotalTestCases = Fn_ExcelGetResultDetail(strExcelPath, "NoTestCases", 1)

	' Pass test cases in batch
	sBatchPassTestCases = Fn_ExcelGetResultDetail(strExcelPath, "PassCount", 1)

	' Fail test cases in batch
	sBatchFailTestCases = Fn_ExcelGetResultDetail(strExcelPath, "FailCount", 1)

	sMailBody = ""
	sMailBody = ("--------------------------------------------------------------------------") & vbcrlf _
	 & ("Batch Summary") & vbcrlf _
	 & ("--------------------------------------------------------------------------") & vbcrlf _
	 & ("Batch Name" & Fn_tabFunction(4) & "|" & vbTab & Fn_GetBatchName(Environment("BatchFldName"))) & vbcrlf _
	 & ("Batch Result" & Fn_tabFunction(3) & "|" & vbTab & sBatchResult) & vbcrlf _
	 & ("Batch Date" & Fn_tabFunction(4) & "|" & vbTab & sBatchDate) & vbcrlf _
	 & ("Batch Executed By" & Fn_tabFunction(3) & "|" & vbTab & Fn_GetEnvValue("user", "AutoUser")) & vbcrlf _
	 & ("Batch Duration" & Fn_tabFunction(3) & "|" & vbTab & sBatchDuration) & vbcrlf _
	 & ("Batch Logs" & Fn_tabFunction(4) & "|" & vbTab & "\\" & sBatchLogFolder) & vbcrlf _
	 & ("Number of Test Cases in Batch" & Fn_tabFunction(1) & "|" & vbTab & sBatchTotalTestCases) & vbcrlf _
	 & ("Number of Passed Test Cases" & Fn_tabFunction(1) & "|" & vbTab & sBatchPassTestCases) & vbcrlf _
	 & ("Number of Failed Test Cases" & Fn_tabFunction(1) & "|" & vbTab & sBatchFailTestCases) & vbcrlf _
	 & ("--------------------------------------------------------------------------") & vbcrlf _
	 & ("Temcenter Server Details") & vbcrlf _
	 & ("--------------------------------------------------------------------------") & vbcrlf _
	 & ("Teamcenter Release" & Fn_tabFunction(2) & "|" & vbTab & Environment("TcRelease")) & vbcrlf _
	 & ("Teamcenter Build" & Fn_tabFunction(3) & "|" & vbTab & Environment("TcBuild")) & vbcrlf _
	 & ("Teamcenter Setup Type" & Fn_tabFunction(2) & "|" & vbTab & "4 Tier") & vbcrlf _
	 & ("Teamcenter Server Host" & Fn_tabFunction(2) & "|" & vbTab & Environment("TcServer")) & vbcrlf _
	 & ("Teamcenter Server OS" & Fn_tabFunction(2) & "|" & vbTab & Environment("TcServerOS")) & vbcrlf _
	 & ("Application Server " & Fn_tabFunction(2) & "|" & vbTab & Environment("ApplicationServer")) & vbcrlf _
	 & ("Teamcenter Database Host" & Fn_tabFunction(1) & "|" & vbTab & Environment("TCDBHost")) & vbcrlf _
	 & ("Teamcenter Database OS" & Fn_tabFunction(2) & "|" & vbTab & Environment("TcDBServerOS")) & vbcrlf _
	 & ("Teamcenter Database Type" & Fn_tabFunction(1) & "|" & vbTab & Environment("DatabaseType") & " " &Environment("DatabaseVersion")) & vbcrlf _
	 & ("--------------------------------------------------------------------------") & vbcrlf _
	 & ("Teamcenter RAC Details") & vbcrlf _
	 & ("--------------------------------------------------------------------------") & vbcrlf _
	 & ("Teamcenter RAC Host" & Fn_tabFunction(2) & "|" & vbTab & Environment("LocalHostName")) & vbcrlf _
	 & ("Teamcenter RAC OS" & Fn_tabFunction(3) & "|" & vbTab & Environment("OS")) & vbcrlf _
	 & ("--------------------------------------------------------------------------") & vbcrlf _
	 & (vbcrlf & vbcrlf & "Thanks & Regards" & vbcrlf & "Automation Team" & vbcrlf & "Teamcenter PV")

	 Fn_EmailBodyTxt = sMailBody

End Function

'**********************************************************************************
' Tab function
'**********************************************************************************
Function Fn_tabFunction(tabCount)
	Dim sTabOpr , i
	For i = 0 to tabCount
		sTabOpr = sTabOpr & vbTab
	Next
	Fn_tabFunction = sTabOpr
End Function

'*********************************************************	Function List		***********************************************************************
'2. Fn_EmailBodyTxt(strExcelPath)
'*********************************************************	Function List		***********************************************************************
Function Fn_EmailBodyHTML(strExcelPath)

	Dim bReturn
	Dim sBatchOwner
	Dim sBatchName
	Dim sBatchResult
	Dim sBatchDate
	Dim sBatchDuration
	Dim sBatchLogFolder
	Dim sBatchTotalTestCases
	Dim sBatchPassTestCases
	Dim sBatchFailTestCases
	Dim sBrowserUsed
	Dim sBatchArea
	Dim sMailBody
	Dim sBatchResultSavedTestCases
	Dim sBatchNotUploadedTestCases

	' Batch Owner
	If Environment("UserName") = "ntpriv" OR lcase(Environment("UserName")) = "administrator"  OR lcase(Environment("UserName")) = "yytpvcad" OR lcase(Environment("UserName")) = "yytpvsad"  Then
		sBatchOwner = Fn_GetEnvValue("user", "AutoUser")
	Else
		sBatchOwner = Fn_GetUserName()
	End If

	' Batch result
	bReturn = Fn_ExcelGetResultDetail(strExcelPath,"FailExist", 1)
	if bReturn = "True" Then
		'sBatchResult = "FAIL" 
		sBatchResult = "<tr><td><font face=""Arial"" SIZE=2>Batch Test Result</td>" &_
					   "<td BGCOLOR=""#FF7373"" align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;FAIL</B></td></tr>"
	else
		'sBatchResult = "PASS"
		sBatchResult = "<tr><td><font face=""Arial"" SIZE=2>Batch Test Result</td>" &_
					   "<td BGCOLOR=""#7DDBA9"" align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;PASS</B></td></tr>"
	end if	

	' Batch Name
	sBatchName = Fn_GetBatchName(Environment("BatchFldName"))

	' Batch date
	sBatchDate = Fn_BatchDate(strExcelPath)

	' Batch duration
	sBatchDuration = Fn_BatchDuration(strExcelPath)

	' Batch log folder
	sBatchLogFolder = Environment.Value("LocalHostName") & "\" & sBatchName

	' Cleanup batch result 
	bReturn = Fn_BatchResultCleanup(strExcelPath, 1)

	' Fetch batch Area
	sBatchArea = Fn_GetTestArea(strExcelPath, 1)

	' Total testcases in batch
	sBatchTotalTestCases = Fn_ExcelGetResultDetail(strExcelPath, "NoTestCases", 1)

	' Pass test cases in batch
	sBatchPassTestCases = Fn_ExcelGetResultDetail(strExcelPath, "PassCount", 1)
	
	'Result Saved test cases in batch
	sBatchResultSavedTestCases= Fn_ExcelGetQARTStatusDetail(strExcelPath, "Results Saved", 1)
	
	'Not Uploaded test cases in batch
	sBatchNotUploadedTestCases= Fn_ExcelGetQARTStatusDetail(strExcelPath, "Not Uploaded", 1)
	

	' Browser Used
	If Instr(Environment("WebBrowserName"),"IE") > 0 Then
		sBrowserUsed = "Internet Explorer"
	ElseIf Instr(Environment("WebBrowserName"),"FF") > 0 Then
		sBrowserUsed = "Firefox" 
	End If

	' Fail test cases in batch
	sBatchFailTestCases = sBatchTotalTestCases - sBatchPassTestCases

	sMailBody = ""
	sMailBody = "<HTML><HEAD></HEAD>" & _
				"<BODY>" & _
				"<table border=1 cellpadding=1 cellspacing=1 bordercolor=""gray"" width=100%>" & _
				"<tr><th colspan=2 bgcolor=""#A0CFEC""><font face=""Arial"" size=4>Teamcenter Automation Test Result</th></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Test Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Batch Test Area</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & Trim(sBatchArea) & "</td></B></tr>" & _
				sBatchResult & _
				"<tr><td><font face=""Arial"" SIZE=2>Total Test Cases in Batch</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchTotalTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Number of Passed Test Cases</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchPassTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Number of Failed Test Cases</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchFailTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Number of Results Uploaded to QART</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchResultSavedTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Number of Results Not Uploaded to QART</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchNotUploadedTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Logs</td><td align=""left"" valign=""center"">" & _
				"<a href=""\\"& sBatchLogFolder &"""><font face=""Arial"" SIZE=2><B>\\"& sBatchLogFolder &"</B></a></td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Batch Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Batch Name</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchName & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Batch Date</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchDate & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Batch Executed By</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchOwner & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Batch Duration</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;<font face=""Arial"" SIZE=2>" & sBatchDuration & "</td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Teamcenter Server Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Release</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcRelease") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Build</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcBuild") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Server Setup Type</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;4 Tier</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Server Host</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcServer") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Server OS</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcServerOS") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Application Server Type</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("ApplicationServer") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Application Host</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcServer") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Application OS</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcServerOS") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Database Type</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("DatabaseType") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Database Host</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TCDBHost") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Database OS</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcDBServerOS") & "</td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Teamcenter RAC Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter RAC Host</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("LocalHostName") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter RAC OS</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("OS") & "</td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Teamcenter WebClient Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter WebClient URL</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment("TcWebServer") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Browser Used</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBrowserUsed & "</td></tr>" & _
				"</table>" & _
				"<p><font face=""Arial"" SIZE=2>Thanks</p>" & _
				"<p><font face=""Arial"" SIZE=2>Automation Team" & _
				"<br>Teamcenter Product Validation<br></p>" & _
				"<p><font face=""Arial"" SIZE=1><br><br><br>PS: This is a automated mail, please do not reply</p>" & _
				"</BODY></HTML>"

	Fn_EmailBodyHTML = sMailBody

End Function

'*********************************************************	Function List		***********************************************************************
'2. Fn_GetUserName()
'*********************************************************	Function List		***********************************************************************
Function Fn_GetUserName()

	Dim objRootDSE, objADSysInfo, objUser
	Dim strNamingContext, strUserDN, sUserName

	Set objRootDSE = GetObject("LDAP://RootDSE") 
	
	If Err.Number = 0 Then 
	    strNamingContext = objRootDSE.Get("defaultNamingContext")  
	Else 
		Fn_GetUserName = False
		Exit Function
	End If 	
	
	Set objADSysInfo = CreateObject("ADSystemInfo")
	strUserDN = objADSysInfo.username 
	Set objUser = Getobject("LDAP://" & strUserDN)
	sUserName = objUser.Get("givenName")  & " " & objUser.Get("sn") 
	
	if 	sUserName <> "" then    	
		Fn_GetUserName =   sUserName
	else
		Fn_GetUserName = False
	end if
	
	Set objUser = Nothing
	Set objADSysInfo = Nothing
	Set objRootDSE = Nothing	
	
End Function
