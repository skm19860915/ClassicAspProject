<%
	Response.Write("Timout: " & Session.Timeout & "<br>")
	Response.Write("<P>SESSION VARIABLES: There are "& Session.Contents.Count & "<br>")
	Response.Write("---------------------<br>")
	   Dim strName, iLoop
	   For Each strName in Session.Contents
		 If IsArray(Session(strName)) then
		   For iLoop = LBound(Session(strName)) to UBound(Session(strName))
			  Response.Write  "   <b>" & strName & "(" & iLoop & ")</b> - " & Session(strName)(iLoop) & "<BR>"
		   Next
		 Else
		   Response.Write "   <b>" & strName & "</b> - " & Session.Contents(strName) & "<BR>"
		 End If
	Next

	dim x,y
	for each x in Request.Cookies
	  response.write("<p>")
	  if Request.Cookies(x).HasKeys then
		for each y in Request.Cookies(x)
		  response.write(x & ":" & y & "=" & Request.Cookies(x)(y))
		  response.write("<br />")
		next
	  else
		Response.Write(x & "=" & Request.Cookies(x) & "<br />")
	  end if
	  response.write "</p>"
	next

	Response.write("<P>Query Strings: <BR>")
	Response.write("-------------------<BR>")
	for each x in Request.QueryString
		Response.write("<P>")
		Response.Write x & " - " & Request.QueryString(x) & "<BR>"
	next
%>
<br><br>
<TABLE>
      <TR>
           <TD>
                <B>Server Varriable</B>
           </TD>
           <TD>
                <B>Value</B>
           </TD>
      </TR>

      <% For Each name In Request.ServerVariables %>
      <TR>
           <TD>
                <%= name %>
           </TD>
           <TD>
                <%= Request.ServerVariables(name) %>
           </TD>
      </TR>
      <% Next %>
</TABLE>

%>