<%
'BindEvents Method @1-3261B678
Sub BindEvents()
    Set Login.Button_DoLogin.CCSEvents("OnClick") = GetRef("Login_Button_DoLogin_OnClick")
End Sub
'End BindEvents Method

Function Login_Button_DoLogin_OnClick(Sender) 'Login_Button_DoLogin_OnClick @3-57DDB0BB

'Login @4-3D04C4CB
    With Login
        If NOT CCLoginUser(.login.Value, .password.Value) Then
            .Errors.addError(CCSLocales.GetText("CCS_LoginError", Empty))
            Login_Button_DoLogin_OnClick = False
            .password.Value = ""
        Else
            If Not IsEmpty(CCGetParam("ret_link", Empty)) Then _
                Redirect = CCGetParam("ret_link", Empty)
            Login_Button_DoLogin_OnClick = True
        End If
    End With
'End Login

End Function 'Close Login_Button_DoLogin_OnClick @3-54C34B28


%>
