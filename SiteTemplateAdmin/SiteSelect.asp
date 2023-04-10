<!-- #INCLUDE FILE="i_Connection.asp" -->
<html>
<head>
<title>SiteSelect</title>
<%
	if Request.Form("SiteID") <> "" Then
		If Request.Form("SiteID") <> "-1" Then
			Session("SiteID") = Request.Form("SiteID")
		Else
			Session("SiteID") = ""
		End	If
		%>
		<script language="Javascript">
			parent.location = "AdminPageTemplateList.asp";
		</script>
		<%
	end if
%>
<style type="text/css">
html, body, form {
	padding: 0px;
	margin: 0px;
	background-color: transparent;
	border: none;
}
</style>
</head>
<body>

<form method="post">
	<span style="font-size:8pt">Select Site:</span>	
	<select name="SiteID" class="KeylexStyleSelect" style="display:inline;" onchange="document.getElementById('SelectButton').disabled = true; this.form.submit();">
		<option value="-1">Add New Site</option>
	<%
	dim rsList
	set rsList = objConn.execute("SELECT SiteID, SiteName FROM Site ORDER BY SiteName")

	do while not rsList.eof
	%>
		<option <% if CStr(Session("SiteID")) = CStr(rsList("SiteID")) then Response.write "selected" %> value="<%=rsList("SiteID")%>"><%=rsList("SiteName")%> (<%=rsList("SiteID")%>)</option>
	<%
		rsList.MoveNext
	loop
	%>
	</select>
	<input type="submit" name="Submit" id="SelectButton" value="Select" class="KeylexStyleButton">
</form>
</body>
</html>