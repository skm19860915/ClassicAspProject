<%
'BindEvents Method @1-BEED6C29
Sub BindEvents()
    Set PageTemplate.SiteMainTemplateID.CCSEvents("BeforeShow") = GetRef("PageTemplate_SiteMainTemplateID_BeforeShow")
    Set PageTemplate.PageTemplateUserLastUpdateBy.CCSEvents("BeforeShow") = GetRef("PageTemplate_PageTemplateUserLastUpdateBy_BeforeShow")
    Set PageTemplate.Blocks.CCSEvents("BeforeShow") = GetRef("PageTemplate_Blocks_BeforeShow")
End Sub
'End BindEvents Method

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


%>
