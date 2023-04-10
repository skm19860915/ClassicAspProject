<%
'BindEvents Method @1-BBD9B291
Sub BindEvents()
    Set PageTemplateSection.PageTemplateSection_TotalRecords.CCSEvents("BeforeShow") = GetRef("PageTemplateSection_PageTemplateSection_TotalRecords_BeforeShow")
End Sub
'End BindEvents Method

Function PageTemplateSection_PageTemplateSection_TotalRecords_BeforeShow(Sender) 'PageTemplateSection_PageTemplateSection_TotalRecords_BeforeShow @38-09DA259F

'Retrieve number of records @39-6E1958F0
    PageTemplateSection.PageTemplateSection_TotalRecords.Value = PageTemplateSection.DataSource.Recordset.RecordCount
'End Retrieve number of records

End Function 'Close PageTemplateSection_PageTemplateSection_TotalRecords_BeforeShow @38-54C34B28


%>
