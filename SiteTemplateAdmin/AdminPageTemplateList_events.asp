<%
'BindEvents Method @1-466DE5C6
Sub BindEvents()
    Set PageTemplate.DataSource.CCSEvents("BeforeExecuteSelect") = GetRef("PageTemplate_DataSource_BeforeExecuteSelect")
End Sub
'End BindEvents Method

Function PageTemplate_DataSource_BeforeExecuteSelect(Sender) 'PageTemplate_DataSource_BeforeExecuteSelect @21-689BEAD2

'Custom Code @54-73254650
' -------------------------
response.Write PageTemplate.DataSource.SQL
response.end
' -------------------------
'End Custom Code

End Function 'Close PageTemplate_DataSource_BeforeExecuteSelect @21-54C34B28


%>
