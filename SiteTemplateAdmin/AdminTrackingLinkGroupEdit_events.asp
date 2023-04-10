<%
'BindEvents Method @1-6018841A
Sub BindEvents()
    Set TrackingLinkCode.TrackingLinkCode_TotalRecords.CCSEvents("BeforeShow") = GetRef("TrackingLinkCode_TrackingLinkCode_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function TrackingLinkCode_TrackingLinkCode_TotalRecords_BeforeShow(Sender) 'TrackingLinkCode_TrackingLinkCode_TotalRecords_BeforeShow @5-F2D83A91

'Retrieve number of records @14-66E47814
    TrackingLinkCode.TrackingLinkCode_TotalRecords.Value = TrackingLinkCode.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close TrackingLinkCode_TrackingLinkCode_TotalRecords_BeforeShow @5-54C34B28


%>
