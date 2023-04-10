<%
'BindEvents Method @1-AD6075C0
Sub BindEvents()
    Set Subscription.Subscription_TotalRecords.CCSEvents("BeforeShow") = GetRef("Subscription_Subscription_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function Subscription_Subscription_TotalRecords_BeforeShow(Sender) 'Subscription_Subscription_TotalRecords_BeforeShow @38-7166143E

'Retrieve number of records @39-3B4FC363
    Subscription.Subscription_TotalRecords.Value = Subscription.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close Subscription_Subscription_TotalRecords_BeforeShow @38-54C34B28


%>
