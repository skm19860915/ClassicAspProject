<%
'BindEvents Method @1-F96ED887
Sub BindEvents()
    Set UserTable.CCSEvents("BeforeInsert") = GetRef("UserTable_BeforeInsert")
    Set UserTable.CCSEvents("BeforeUpdate") = GetRef("UserTable_BeforeUpdate")
End Sub
'End BindEvents Method

Function UserTable_BeforeInsert(Sender) 'UserTable_BeforeInsert @56-FF0E3748

'Custom Code @70-73254650
' -------------------------
response.Write usertable.command.sql
response.end
' -------------------------
'End Custom Code

End Function 'Close UserTable_BeforeInsert @56-54C34B28

Function UserTable_BeforeUpdate(Sender) 'UserTable_BeforeUpdate @56-835461E7

'Custom Code @71-73254650
' -------------------------
response.Write UserTable.command
response.end
' -------------------------
'End Custom Code

End Function 'Close UserTable_BeforeUpdate @56-54C34B28


%>
