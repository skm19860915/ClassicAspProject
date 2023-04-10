<%
' --------------------------------------------------------------------
' Filename     : i_Connection.asp
' Purpose      : Include Connection Object to DB
' Date Created : 6/28/2006
' Created By   : Ben Shimshak
' Updated On   : 
' Required     : None
'
' Functions    : None
' --------------------------------------------------------------------
%>
<%
	Server.ScriptTimeout = 60
	Dim objConn,DBServer,DBDatabase,DBLogin,DBPassword
	Set objConn = Server.CreateObject("ADODB.Connection")
	
	DBServer   = "localhost"
	DBDatabase = "SiteTemplate"
	DBLogin    = "keylex"
	DBPassword = "KEY!@#"
	objConn.Open ("Provider=SQLOLEDB.1;User ID=" & DBLogin& ";Password=" & DBPassword & ";Persist Security Info=True;Initial Catalog=" & DBDatabase & ";Data Source=" & DBServer)
%>