<%
'BindEvents Method @1-B3EBA499
Sub BindEvents()
    Set EmailTemplate.ArchivedRows.CCSEvents("BeforeShow") = GetRef("EmailTemplate_ArchivedRows_BeforeShow")
    Set EmailTemplate.EmailTemplateUserLastUpdateBy.CCSEvents("BeforeShow") = GetRef("EmailTemplate_EmailTemplateUserLastUpdateBy_BeforeShow")
    Set EmailTemplate.DataSource.CCSEvents("BeforeBuildInsert") = GetRef("EmailTemplate_DataSource_BeforeBuildInsert")
    Set EmailTemplate.CCSEvents("BeforeUpdate") = GetRef("EmailTemplate_BeforeUpdate")
    Set EmailTemplate.DataSource.CCSEvents("BeforeExecuteUpdate") = GetRef("EmailTemplate_DataSource_BeforeExecuteUpdate")
    Set EmailTemplate.DataSource.CCSEvents("BeforeExecuteDelete") = GetRef("EmailTemplate_DataSource_BeforeExecuteDelete")
End Sub
'End BindEvents Method

Function EmailTemplate_ArchivedRows_BeforeShow(Sender) 'EmailTemplate_ArchivedRows_BeforeShow @82-E6BFC39B

'Custom Code @83-73254650
' -------------------------
EmailTemplate.ArchivedRows.Value = CCDLookup("COUNT(*)", "EmailTemplateArchive", "EmailTemplateSiteID = " & Session("SiteID") & " AND EmailTemplateID = '" & Request.QueryString("EmailTemplateID") & "'", DBSystem)
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_ArchivedRows_BeforeShow @82-54C34B28

Function EmailTemplate_EmailTemplateUserLastUpdateBy_BeforeShow(Sender) 'EmailTemplate_EmailTemplateUserLastUpdateBy_BeforeShow @38-06BD5505

'Custom Code @39-73254650
' -------------------------
If Session("UserLogin") = "" Then response.redirect("/")
EmailTemplate.EmailTemplateUserLastUpdateBy.Value = Session("UserLogin")
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_EmailTemplateUserLastUpdateBy_BeforeShow @38-54C34B28

Function EmailTemplate_DataSource_BeforeBuildInsert(Sender) 'EmailTemplate_DataSource_BeforeBuildInsert @2-5C6D6621

'Custom Code @26-73254650
' -------------------------
	If EmailTemplate.EmailTemplateImagePhysicalPath.Value <> "" Then
		EmailTemplate.EmailTemplateImagePhysicalPath.Value = Server.MapPath(EmailTemplate.EmailTemplateImagePath.value)
	End If
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_DataSource_BeforeBuildInsert @2-54C34B28

Function EmailTemplate_BeforeUpdate(Sender) 'EmailTemplate_BeforeUpdate @2-0FE86081

'Custom Code @27-73254650
' -------------------------
	If EmailTemplate.EmailTemplateImagePhysicalPath.Value <> "" Then
		EmailTemplate.EmailTemplateImagePhysicalPath.Value = Server.MapPath(EmailTemplate.EmailTemplateImagePath.value)
	End If
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_BeforeUpdate @2-54C34B28

Function EmailTemplate_DataSource_BeforeExecuteUpdate(Sender) 'EmailTemplate_DataSource_BeforeExecuteUpdate @2-DC268FD5

'Custom Code @40-73254650
' -------------------------
EmailTemplate.Command.SQL = "INSERT INTO EmailTemplateArchive SELECT * FROM EmailTemplate WHERE EmailTemplateID = " & Request.QueryString("EmailTemplateID") & " " & EmailTemplate.Command.SQL
EmailTemplate.Command.SQL = EmailTemplate.Command.SQL & " UPDATE EmailTemplate SET EmailTemplateUserLastUpdateDateTime = getdate() WHERE EmailTemplateID = " & Request.QueryString("EmailTemplateID")
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_DataSource_BeforeExecuteUpdate @2-54C34B28

Function EmailTemplate_DataSource_BeforeExecuteDelete(Sender) 'EmailTemplate_DataSource_BeforeExecuteDelete @2-3866427A

'Custom Code @41-73254650
' -------------------------
EmailTemplate.Command.SQL = "INSERT INTO EmailTemplateArchive SELECT * FROM EmailTemplate WHERE EmailTemplateID = " & Request.QueryString("EmailTemplateID") & " " & EmailTemplate.Command.SQL
' -------------------------
'End Custom Code

End Function 'Close EmailTemplate_DataSource_BeforeExecuteDelete @2-54C34B28


%>
