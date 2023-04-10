<%
'BindEvents Method @1-2A0313FD
Sub BindEvents()
    Set EmailTemplate.EmailTemplate_TotalRecords.CCSEvents("BeforeShow") = GetRef("EmailTemplate_EmailTemplate_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function EmailTemplate_EmailTemplate_TotalRecords_BeforeShow(Sender) 'EmailTemplate_EmailTemplate_TotalRecords_BeforeShow @4-A2ADBC14

'Retrieve number of records @24-43E9BF49
    EmailTemplate.EmailTemplate_TotalRecords.Value = EmailTemplate.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close EmailTemplate_EmailTemplate_TotalRecords_BeforeShow @4-54C34B28


%>
