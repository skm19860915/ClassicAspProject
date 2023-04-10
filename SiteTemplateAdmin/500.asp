<%
Set o = Server.getLastError()
Response.Write "<p>AspCode: " & o.AspCode & "</p>"
Response.Write "<p>Category: " & o.Category & "</p>"
Response.Write "<p>File: " & o.File  & "</p>"
Response.Write "<p>Description: " & o.Description & "</p>"
Response.Write "<p>Source: " & Server.HtmlEncode(o.Source) & "</p>"
Response.Write "<p>Line: " & o.Line & "</p>"
Response.Write "<p>AspDescription: " & o.AspDescription & "</p>"
Set o = Nothing


%>
