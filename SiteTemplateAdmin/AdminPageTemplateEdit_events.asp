<%
'BindEvents Method @1-B342C76E
Sub BindEvents()
    Set PageTemplate.ArchivedRows.CCSEvents("BeforeShow") = GetRef("PageTemplate_ArchivedRows_BeforeShow")
    Set PageTemplate.SiteMainTemplateID.CCSEvents("BeforeShow") = GetRef("PageTemplate_SiteMainTemplateID_BeforeShow")
    Set PageTemplate.PageTemplateUserLastUpdateBy.CCSEvents("BeforeShow") = GetRef("PageTemplate_PageTemplateUserLastUpdateBy_BeforeShow")
    Set PageTemplate.Blocks.CCSEvents("BeforeShow") = GetRef("PageTemplate_Blocks_BeforeShow")
    Set PageTemplate.DataSource.CCSEvents("BeforeExecuteUpdate") = GetRef("PageTemplate_DataSource_BeforeExecuteUpdate")
    Set PageTemplate.DataSource.CCSEvents("BeforeExecuteDelete") = GetRef("PageTemplate_DataSource_BeforeExecuteDelete")
End Sub
'End BindEvents Method

Function PageTemplate_ArchivedRows_BeforeShow(Sender) 'PageTemplate_ArchivedRows_BeforeShow @82-9D60B53B

'Custom Code @83-73254650
' -------------------------
Dim SiteID, SearchString, ArchivedRS, intArchivedRows

SiteID = Session("SiteID")
If SiteID = "" Then response.redirect("/")
SearchString = PageTemplate.PageTemplateNickname.Value
intArchivedRows = 0

If Len(SearchString) > 0 Then
	SearchString = "'" & Replace(SearchString, "'", "''") & "'"

	Set ArchivedRS = DBSystem.Execute("EXEC sp_getPageTemplateList " & SiteID & ", " & SearchString & ", 'PageTemplateArchive', 'PageTemplateID'")

	While Not ArchivedRS.EOF
		intArchivedRows = intArchivedRows + 1
		ArchivedRS.MoveNext
	Wend

End If

PageTemplate.ArchivedRows.Value = intArchivedRows
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_ArchivedRows_BeforeShow @82-54C34B28

Function PageTemplate_SiteMainTemplateID_BeforeShow(Sender) 'PageTemplate_SiteMainTemplateID_BeforeShow @66-D1DC7E6D

'Custom Code @67-73254650
' -------------------------
PageTemplate.SiteMainTemplateID.Value = CCDLookup("SiteMainTemplateID", "Site", "SiteID = " & Session("SiteID"), DBSystem)
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_SiteMainTemplateID_BeforeShow @66-54C34B28

Function PageTemplate_PageTemplateUserLastUpdateBy_BeforeShow(Sender) 'PageTemplate_PageTemplateUserLastUpdateBy_BeforeShow @73-78756F0F

'Custom Code @74-73254650
' -------------------------
PageTemplate.PageTemplateUserLastUpdateBy.Value = Session("UserLogin")
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_PageTemplateUserLastUpdateBy_BeforeShow @73-54C34B28

Function PageTemplate_Blocks_BeforeShow(Sender) 'PageTemplate_Blocks_BeforeShow @63-6F3D29D8

'Custom Code @64-73254650
' -------------------------
Dim maxBlocks, rsOptions, optionList, i
maxBlocks = CCDLookup("SiteMaxTemplateBlocks", "Site", "SiteID = " & Session("SiteID"), DBSystem)

Set rsOptions = DBSystem.Execute("SELECT PageTemplateNickname, PageTemplateName FROM PageTemplate WHERE PageTemplatePageType = 'Block' AND PageTemplateSiteID = '" & Session("SiteID") & "'")

optionList = vbCrLf & "<option value=""""></option>" & vbCrLf

While Not rsOptions.EOF
	optionList = optionList & "<option value=""" & rsOptions("PageTemplateNickname") & """>" & rsOptions("PageTemplateName") & "</option>" & vbCrLf

	rsOptions.MoveNext
Wend

For i = 1 To maxBlocks
	PageTemplate.Blocks.Value = PageTemplate.Blocks.Value &_
		"Block #" & i & vbCrLf &_
		"<select id=""PageTemplateBlock" & i & """>" & optionList & "</select>" & vbCrLf & "<br />" & vbCrLf
Next
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_Blocks_BeforeShow @63-54C34B28

Function PageTemplate_DataSource_BeforeExecuteUpdate(Sender) 'PageTemplate_DataSource_BeforeExecuteUpdate @2-F16D7D72

'Custom Code @75-73254650
' -------------------------
PageTemplate.Command.SQL = "INSERT INTO PageTemplateArchive SELECT * FROM PageTemplate WHERE PageTemplateID = " & Request.QueryString("PageTemplateID") & " " & PageTemplate.Command.SQL
PageTemplate.Command.SQL = PageTemplate.Command.SQL & " UPDATE PageTemplate SET PageTemplateUserLastUpdateDateTime = getdate() WHERE PageTemplateID = " & Request.QueryString("PageTemplateID")
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_DataSource_BeforeExecuteUpdate @2-54C34B28

Function PageTemplate_DataSource_BeforeExecuteDelete(Sender) 'PageTemplate_DataSource_BeforeExecuteDelete @2-152DB0DD

'Custom Code @76-73254650
' -------------------------
PageTemplate.Command.SQL = "INSERT INTO PageTemplateArchive SELECT * FROM PageTemplate WHERE PageTemplateID = " & Request.QueryString("PageTemplateID") & " " & PageTemplate.Command.SQL
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_DataSource_BeforeExecuteDelete @2-54C34B28


%>
