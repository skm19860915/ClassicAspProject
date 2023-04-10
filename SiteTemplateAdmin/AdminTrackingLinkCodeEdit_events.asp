<%
'BindEvents Method @1-D3220815
Sub BindEvents()
    Set TrackingLinkCode.TrackingLinkCode.CCSEvents("BeforeShow") = GetRef("TrackingLinkCode_TrackingLinkCode_BeforeShow")
    Set TrackingLinkCode.TrackingLinkCodeLink.CCSEvents("BeforeShow") = GetRef("TrackingLinkCode_TrackingLinkCodeLink_BeforeShow")
End Sub
'End BindEvents Method

Function TrackingLinkCode_TrackingLinkCode_BeforeShow(Sender) 'TrackingLinkCode_TrackingLinkCode_BeforeShow @11-C2D465C9

'Custom Code @16-73254650
' -------------------------
if not TrackingLinkCode.EditMode then
	Dim oldTimeOut
	oldTimeOut =  Server.ScriptTimeout
	Server.ScriptTimeout = 20
	dim rsObj, strCode, duplicateCount
	randomize
	strCode = UCASE(getrandomcode(CCDLookup("SiteTrackingLinkLength", "Site", "SiteID = " & Session("SiteID"), DBSystem),1))
	'response.write strCode & "<br>"
	set rsObj = DBSystem.Execute("SELECT Count(*) as KeyCount FROM TrackingLinkCode WHERE TrackingLinkCode='" & strCode & "'")
	duplicateCount = rsObj("KeyCount")
	if duplicateCount = 0 then
		TrackingLinkCode.TrackingLinkCode.Value = strCode
	else
		do while duplicateCount <> 0 
			randomize
			strCode = UCASE(getrandomcode(CCDLookup("SiteTrackingLinkLength", "Site", "SiteID = " & Session("SiteID"), DBSystem),1))
			'response.write strCode & "<br>"
			set rsObj = DBSystem.Execute("SELECT Count(*) as KeyCount FROM TrackingLinkCode WHERE TrackingLinkCode='" & strCode & "'")
			duplicateCount = rsObj("KeyCount")
		loop
		TrackingLinkCode.TrackingLinkCode.Value = strCode
	end if
	Server.ScriptTimeout = oldTimeOut
end if
' -------------------------
'End Custom Code

End Function 'Close TrackingLinkCode_TrackingLinkCode_BeforeShow @11-54C34B28

Function TrackingLinkCode_TrackingLinkCodeLink_BeforeShow(Sender) 'TrackingLinkCode_TrackingLinkCodeLink_BeforeShow @17-9D6D081F

'Custom Code @18-73254650
' -------------------------
TrackingLinkCode.TrackingLinkCodeLink.value = CCDLookup("SiteDefaultURL + '?' + SiteTrackingString + '='", "Site", "SiteID = " & Session("SiteID"),DBSystem) & TrackingLinkCode.TrackingLinkCode.value
' -------------------------
'End Custom Code

End Function 'Close TrackingLinkCode_TrackingLinkCodeLink_BeforeShow @17-54C34B28


%>
