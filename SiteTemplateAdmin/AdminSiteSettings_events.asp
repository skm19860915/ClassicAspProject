<%
'BindEvents Method @1-9C137ADF
Sub BindEvents()
    Set Site.DebugDomainName.CCSEvents("BeforeShow") = GetRef("Site_DebugDomainName_BeforeShow")
    Set Site.CCSEvents("AfterInsert") = GetRef("Site_AfterInsert")
End Sub
'End BindEvents Method


Function Site_DebugDomainName_BeforeShow(Sender) 'Site_DebugDomainName_BeforeShow @97-F8108936

'Custom Code @98-73254650
' -------------------------
If Site.SiteDomainName.Value <> "" Then
	Site.DebugDomainName.Value = Split(Site.SiteDomainName.Value, ",")(0)
End If
' -------------------------
'End Custom Code

End Function 'Close Site_DebugDomainName_BeforeShow @97-54C34B28



Function Site_AfterInsert(Sender) 'Site_AfterInsert @55-F4973352

'Custom Code @88-73254650
' -------------------------
	Dim rsRecordSet

	rsRecordSet = DBSystem.execute("SELECT @@IDENTITY AS SiteID")
	Session("SiteID") = rsRecordSet("SiteID")

' -------------------------
'End Custom Code

End Function 'Close Site_AfterInsert @55-54C34B28


%>
