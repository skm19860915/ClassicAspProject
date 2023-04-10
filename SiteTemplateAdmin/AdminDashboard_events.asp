<%
'BindEvents Method @1-A81A3B5D
Sub BindEvents()
    Set Label1.CCSEvents("BeforeShow") = GetRef("Label1_BeforeShow")
End Sub
'End BindEvents Method

Function Label1_BeforeShow(Sender) 'Label1_BeforeShow @35-88342291

'Custom Code @36-73254650
' -------------------------
	'send the data
	Dim objXMLHTTP
	Set objXMLHTTP = Server.CreateObject("MSXML2.XMLHTTP")
	response.Write "http://" & Request.ServerVariables("HTTP_HOST") & "/AdminDashboardStats.asp"
	objXMLHTTP.Open "GET", "http://" & Request.ServerVariables("HTTP_HOST") & "/AdminDashboardStats.asp", false
	objXMLHTTP.Send ""

	Label1.Value objXMLHTTP.responseText

	Set objXMLHTTP = nothing

' -------------------------
'End Custom Code

End Function 'Close Label1_BeforeShow @35-54C34B28


%>
