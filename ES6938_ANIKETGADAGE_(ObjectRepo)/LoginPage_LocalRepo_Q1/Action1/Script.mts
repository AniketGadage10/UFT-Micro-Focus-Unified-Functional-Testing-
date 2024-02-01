'Try creating a local OR for flight reservation application login page.

WpfWindow("Micro Focus MyFlight Sample").InsightObject("InsightObject").Click
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set "john"

WpfWindow("Micro Focus MyFlight Sample").InsightObject("InsightObject_2").Click

WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "65b8a27e12557951a935"

WpfWindow("Micro Focus MyFlight Sample").InsightObject("InsightObject_3").Click

