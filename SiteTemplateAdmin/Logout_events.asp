<%
'BindEvents Method @1-90ADDF24
Sub BindEvents()
    Set CCSEvents("AfterInitialize") = GetRef("Page_AfterInitialize")
End Sub
'End BindEvents Method

Function Page_AfterInitialize(Sender) 'Page_AfterInitialize @1-5C791CCC

'Logout @2-00022735
    CCLogoutUser
    Redirect = "AdminDashboard.asp"
'End Logout

End Function 'Close Page_AfterInitialize @1-54C34B28


%>
