 
 Environment.LoadFromFile(Environment.Value("Path")&Environment.Value("TestName")&".xml")
 
Login Environment.Value("AgentName"),Environment.Value("Password")

WpfWindow("Micro Focus MyFlight Sample").Close
