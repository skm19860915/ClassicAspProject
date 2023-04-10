<%
'BindEvents Method @1-5CDFC931
Sub BindEvents()
    Set EmailTemplate.EmailTemplateSection.CCSEvents("BeforeShow") = GetRef("EmailTemplate_EmailTemplateSection_BeforeShow")
    Set EmailTemplate.Alt_EmailTemplateSection.CCSEvents("BeforeShow") = GetRef("EmailTemplate_Alt_EmailTemplateSection_BeforeShow")
    Set EmailTemplate.EmailTemplate_TotalRecords.CCSEvents("BeforeShow") = GetRef("EmailTemplate_EmailTemplate_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function EmailTemplate_EmailTemplateSection_BeforeShow(Sender) 'EmailTemplate_EmailTemplateSection_BeforeShow @52-DCCD9C2F

'Custom Code @54-73254650
' -------------------------
Dim strSection

Select Case EmailTemplate.DataSource.EmailTemplateSection
	Case "0"	strSection = "Client Confirmation"
	Case "1"	strSection = "Client Notification"
	Case "2"	strSection = "Admin Confirmation"
	Case "3"	strSection = "Admin Notification"
End Select

EmailTemplate.EmailTemplateSection.Value = strSection
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_EmailTemplateSection_BeforeShow @52-54C34B28

Function EmailTemplate_Alt_EmailTemplateSection_BeforeShow(Sender) 'EmailTemplate_Alt_EmailTemplateSection_BeforeShow @53-1ED2A14E

'Custom Code @55-73254650
' -------------------------
Dim strSection

Select Case EmailTemplate.DataSource.EmailTemplateSection
	Case "0"	strSection = "Client Confirmation"
	Case "1"	strSection = "Client Notification"
	Case "2"	strSection = "Admin Confirmation"
	Case "3"	strSection = "Admin Notification"
End Select

EmailTemplate.Alt_EmailTemplateSection.Value = strSection
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_Alt_EmailTemplateSection_BeforeShow @53-54C34B28

Function EmailTemplate_EmailTemplate_TotalRecords_BeforeShow(Sender) 'EmailTemplate_EmailTemplate_TotalRecords_BeforeShow @4-A2ADBC14

'Retrieve number of records @24-43E9BF49
    EmailTemplate.EmailTemplate_TotalRecords.Value = EmailTemplate.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close EmailTemplate_EmailTemplate_TotalRecords_BeforeShow @4-54C34B28


%>
