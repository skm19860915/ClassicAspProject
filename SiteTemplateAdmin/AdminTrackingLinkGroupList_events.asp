<%
'BindEvents Method @1-2A07960F
Sub BindEvents()
    Set TrackingLinkGroup.TrackingLinkGroup_TotalRecords.CCSEvents("BeforeShow") = GetRef("TrackingLinkGroup_TrackingLinkGroup_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function TrackingLinkGroup_TrackingLinkGroup_TotalRecords_BeforeShow(Sender) 'TrackingLinkGroup_TrackingLinkGroup_TotalRecords_BeforeShow @9-3F9F2BE5

'Retrieve number of records @10-79B081FD
    TrackingLinkGroup.TrackingLinkGroup_TotalRecords.Value = TrackingLinkGroup.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close TrackingLinkGroup_TrackingLinkGroup_TotalRecords_BeforeShow @9-54C34B28


%>
