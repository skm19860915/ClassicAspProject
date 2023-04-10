<!-- #include file="i_Connection.asp" -->
<%

Dim SiteID
SiteID = Request.QueryString("SiteID")

If SiteID = "" Or Not IsNumeric(SiteID) Then Response.End

Function statExists(rsStats, statName)
	If rsStats.EOF Then
		statExists = False
	Else
		statExists = (rsStats.Fields.Item(0) = statName)
	End If
End Function

Dim strSQL, rsStats, statTotalCurrentSessions, statCurrentSessions, statCurrentSessionArray, statVisitsToday, statVisitsWeek, statVisitsMonth, statVisitsYear, statTopPages, statTopReferrers, statEventLogInfos, statEventLogWarnings, statEventLogAlerts, statASPErrors

strSQL = "EXEC sp_getSiteStats " & SiteID
Set rsStats = objConn.Execute(strSQL)

statTotalCurrentSessions = 0
While statExists(rsStats, "Current Session")
	'192.168.10.21|bigsimpfan|270989788|/page.asp
	statCurrentSessionArray = Split(rsStats.Fields.Item(2), "|")

	statCurrentSessions = statCurrentSessions & "<br/> - " &_
		statCurrentSessionArray(0) & " - " &_
		statCurrentSessionArray(1) & " " &_
		statCurrentSessionArray(2) & " " &_
		"[" & statCurrentSessionArray(3) & "] - " &_
		statCurrentSessionArray(4)

	statTotalCurrentSessions = statTotalCurrentSessions + 1
	rsStats.MoveNext
Wend

If statExists(rsStats, "Visits Today") Then
	statVisitsToday = rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

If statExists(rsStats, "Visits Week") Then
	statVisitsWeek = rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

If statExists(rsStats, "Visits Month") Then
	statVisitsMonth = rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

If statExists(rsStats, "Visits Year") Then
	statVisitsYear = rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

While statExists(rsStats, "Top Page")
	statTopPages = statTopPages & "<br/> - " &_
		rsStats.Fields.Item(1) & " - " &_
		rsStats.Fields.Item(2)

	rsStats.MoveNext
Wend

While statExists(rsStats, "Top Referer")
	statTopReferers = statTopReferers & "<br/> - " &_
		rsStats.Fields.Item(1) & " - " &_
		rsStats.Fields.Item(2)
	rsStats.MoveNext
Wend

If statExists(rsStats, "EventLog - Infos") Then
	statEventLogInfos = rsStats.Fields.Item(2) & " / " & rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

If statExists(rsStats, "EventLog - Warnings") Then
	statEventLogWarnings = rsStats.Fields.Item(2) & " / " & rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

If statExists(rsStats, "EventLog - Alerts") Then
	statEventLogAlerts = rsStats.Fields.Item(2) & " / " & rsStats.Fields.Item(1)

	rsStats.MoveNext
End If

If statExists(rsStats, "ASP Errors") Then
	statASPErrors = rsStats.Fields.Item(1)

	If Not IsNull(rsStats.Fields.Item(2)) Then statASPErrors = statASPErrors & " - " & rsStats.Fields.Item(2)

	rsStats.MoveNext
End If

%>
<font face="verdana" size="1">
<br><br>

<strong>Current Sessions (<%=statTotalCurrentSessions%>)</strong> :
<%=statCurrentSessions%>
<br>
<strong>Visits Today</strong> : <%=statVisitsToday%>
<br>
<strong>Visits Week</strong> : <%=statVisitsWeek%>
<br>
<strong>Visits Month</strong> : <%=statVisitsMonth%>
<br>
<strong>Visits Year</strong> : <%=statVisitsYear%>
<br><br>
<strong>Top 3 Pages</strong>:
<%=statTopPages%>
<br><br>
<strong>Top 3 Referers</strong>:
<%=statTopReferers%>
<br><br>
<strong>EventLogs - Infos</strong> : <%=statEventLogInfos%>
<br>
<strong>EventLogs - Warnings</strong> : <%=statEventLogWarnings%>
<br>
<strong>EventLogs - Alerts</strong> : <%=statEventLogAlerts%>
<br>
<strong>Errors</strong> : <%=statASPErrors%>
