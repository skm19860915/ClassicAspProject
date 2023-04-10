<%
'BindEvents Method @1-F67BB382
Sub BindEvents()
    Set UserTable.UserTable_TotalRecords.CCSEvents("BeforeShow") = GetRef("UserTable_UserTable_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function UserTable_UserTable_TotalRecords_BeforeShow(Sender) 'UserTable_UserTable_TotalRecords_BeforeShow @40-FDFE7348

'Retrieve number of records @41-125D0090
    UserTable.UserTable_TotalRecords.Value = UserTable.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close UserTable_UserTable_TotalRecords_BeforeShow @40-54C34B28


%>
