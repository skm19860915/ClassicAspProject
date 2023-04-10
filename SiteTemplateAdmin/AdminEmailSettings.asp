<%@ CodePage=1252 %>
<%
'Include Common Files @1-C42A9701
%>
<!-- #INCLUDE VIRTUAL="/Common.asp"-->
<!-- #INCLUDE VIRTUAL="/Cache.asp" -->
<!-- #INCLUDE VIRTUAL="/Template.asp" -->
<!-- #INCLUDE VIRTUAL="/Sorter.asp" -->
<!-- #INCLUDE VIRTUAL="/Navigator.asp" -->
<%
'End Include Common Files

'Initialize Page @1-F0EE8708
' Variables
Dim PathToRoot, ScriptPath, TemplateFilePath
Dim FileName
Dim Redirect
Dim Tpl, HTMLTemplate
Dim TemplateFileName
Dim ComponentName
Dim PathToCurrentPage

' Events
Dim CCSEvents
Dim CCSEventResult

' Connections
Dim DBSystem

' Page controls
Dim Menu
Dim EmailDefault
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminEmailSettings.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminEmailSettings.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-3A73A510
CCSecurityRedirect "50", Empty
'End Authenticate User

'Initialize Objects @1-B9BED228
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set EmailDefault = new clsRecordEmailDefault
Menu.Value = DHTMLMenu

EmailDefault.Initialize DBSystem

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-125DDFF5
EmailDefault.Operation
'End Execute Components

'Go to destination page @1-6D35F4FD
If NOT ( Redirect = "" ) Then
    UnloadPage
    Response.Redirect Redirect
End If
'End Go to destination page

'Initialize HTML Template @1-317B9EBC
CCSEventResult = CCRaiseEvent(CCSEvents, "OnInitializeView", Nothing)
Set HTMLTemplate = new clsTemplate
Set HTMLTemplate.Cache = TemplatesRepository
HTMLTemplate.LoadTemplate TemplateFilePath & TemplateFileName
Set Tpl = HTMLTemplate.Block("main")
CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Nothing)
'End Initialize HTML Template

'Show Page @1-2FA5FFEE
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, EmailDefault))
ChildControls.Show
Dim MainHTML
HTMLTemplate.Parse "main", False
MainHTML = HTMLTemplate.GetHTML("main")
CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeOutput", Nothing)
If CCSEventResult Then Response.Write MainHTML
'End Show Page

'Unload Page @1-CB210C62
UnloadPage
Set Tpl = Nothing
Set HTMLTemplate = Nothing
'End Unload Page

'UnloadPage Sub @1-FB242D24
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set EmailDefault = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordEmailDefault 'EmailDefault Class @33-9FEB31E2

'EmailDefault Variables @33-4748F2CE

    ' Public variables
    Public ComponentName
    Public HTMLFormAction
    Public PressedButton
    Public Errors
    Public FormSubmitted
    Public EditMode
    Public Visible
    Public Recordset
    Public TemplateBlock

    Public CCSEvents
    Private CCSEventResult

    Public InsertAllowed
    Public UpdateAllowed
    Public DeleteAllowed
    Public ReadAllowed
    Public DataSource
    Public Command
    Public ValidatingControls
    Public Controls

    ' Class variables
    Dim EmailDefaultSiteURL
    Dim EmailDefaultSMTPServer
    Dim EmailDefaultSMTPPort
    Dim EmailDefaultSMTPAuthentication
    Dim EmailDefaultSMTPUserName
    Dim EmailDefaultSMTPPassword
    Dim EmailDefaultFromAddress
    Dim EmailDefaultFromName
    Dim EmailDefaultReplyToAddress
    Dim EmailDefaultReplyToName
    Dim EmailDefaultBounceBackEmail
    Dim EmailDefaultBounceBackEmailPassword
    Dim EmailDefaultPOPServer
    Dim EmailDefaultPOPPort
    Dim EmailDefaultDNS
    Dim EmailDefaultDNSHelo
    Dim EmailDefaultAdminEmailAddress
    Dim EmailDefaultUnsubscribeURL
    Dim Button_Update
    Dim Button_Cancel
'End EmailDefault Variables

'EmailDefault Class_Initialize Event @33-BD57528E
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsEmailDefaultDataSource
        Set Command = New clsCommand
        InsertAllowed = False
        UpdateAllowed = True
        DeleteAllowed = False
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "EmailDefault")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "EmailDefault"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set EmailDefaultSiteURL = CCCreateControl(ccsTextBox, "EmailDefaultSiteURL", "Site URL", ccsText, Empty, CCGetRequestParam("EmailDefaultSiteURL", Method))
        Set EmailDefaultSMTPServer = CCCreateControl(ccsTextBox, "EmailDefaultSMTPServer", "SMTPServer", ccsText, Empty, CCGetRequestParam("EmailDefaultSMTPServer", Method))
        Set EmailDefaultSMTPPort = CCCreateControl(ccsTextBox, "EmailDefaultSMTPPort", "SMTPPort", ccsInteger, Empty, CCGetRequestParam("EmailDefaultSMTPPort", Method))
        Set EmailDefaultSMTPAuthentication = CCCreateControl(ccsCheckBox, "EmailDefaultSMTPAuthentication", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("EmailDefaultSMTPAuthentication", Method))
        Set EmailDefaultSMTPUserName = CCCreateControl(ccsTextBox, "EmailDefaultSMTPUserName", "SMTPUser Name", ccsText, Empty, CCGetRequestParam("EmailDefaultSMTPUserName", Method))
        Set EmailDefaultSMTPPassword = CCCreateControl(ccsTextBox, "EmailDefaultSMTPPassword", "SMTPPassword", ccsText, Empty, CCGetRequestParam("EmailDefaultSMTPPassword", Method))
        Set EmailDefaultFromAddress = CCCreateControl(ccsTextBox, "EmailDefaultFromAddress", "From Address", ccsText, Empty, CCGetRequestParam("EmailDefaultFromAddress", Method))
        Set EmailDefaultFromName = CCCreateControl(ccsTextBox, "EmailDefaultFromName", "From Name", ccsText, Empty, CCGetRequestParam("EmailDefaultFromName", Method))
        Set EmailDefaultReplyToAddress = CCCreateControl(ccsTextBox, "EmailDefaultReplyToAddress", "Reply To Address", ccsText, Empty, CCGetRequestParam("EmailDefaultReplyToAddress", Method))
        Set EmailDefaultReplyToName = CCCreateControl(ccsTextBox, "EmailDefaultReplyToName", "Reply To Name", ccsText, Empty, CCGetRequestParam("EmailDefaultReplyToName", Method))
        Set EmailDefaultBounceBackEmail = CCCreateControl(ccsTextBox, "EmailDefaultBounceBackEmail", "Bounce Back Email", ccsText, Empty, CCGetRequestParam("EmailDefaultBounceBackEmail", Method))
        Set EmailDefaultBounceBackEmailPassword = CCCreateControl(ccsTextBox, "EmailDefaultBounceBackEmailPassword", "Bounce Back Email Password", ccsText, Empty, CCGetRequestParam("EmailDefaultBounceBackEmailPassword", Method))
        Set EmailDefaultPOPServer = CCCreateControl(ccsTextBox, "EmailDefaultPOPServer", "POPServer", ccsText, Empty, CCGetRequestParam("EmailDefaultPOPServer", Method))
        Set EmailDefaultPOPPort = CCCreateControl(ccsTextBox, "EmailDefaultPOPPort", "POPPort", ccsInteger, Empty, CCGetRequestParam("EmailDefaultPOPPort", Method))
        Set EmailDefaultDNS = CCCreateControl(ccsCheckBox, "EmailDefaultDNS", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("EmailDefaultDNS", Method))
        Set EmailDefaultDNSHelo = CCCreateControl(ccsTextBox, "EmailDefaultDNSHelo", "DNSHelo", ccsText, Empty, CCGetRequestParam("EmailDefaultDNSHelo", Method))
        Set EmailDefaultAdminEmailAddress = CCCreateControl(ccsTextBox, "EmailDefaultAdminEmailAddress", "Admin Email Address", ccsText, Empty, CCGetRequestParam("EmailDefaultAdminEmailAddress", Method))
        Set EmailDefaultUnsubscribeURL = CCCreateControl(ccsTextBox, "EmailDefaultUnsubscribeURL", "Unsubscribe URL", ccsText, Empty, CCGetRequestParam("EmailDefaultUnsubscribeURL", Method))
        Set Button_Update = CCCreateButton("Button_Update", Method)
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(EmailDefaultSiteURL, EmailDefaultSMTPServer, EmailDefaultSMTPPort, EmailDefaultSMTPAuthentication, EmailDefaultSMTPUserName, EmailDefaultSMTPPassword, EmailDefaultFromAddress, _
             EmailDefaultFromName, EmailDefaultReplyToAddress, EmailDefaultReplyToName, EmailDefaultBounceBackEmail, EmailDefaultBounceBackEmailPassword, EmailDefaultPOPServer, EmailDefaultPOPPort, EmailDefaultDNS, _
             EmailDefaultDNSHelo, EmailDefaultAdminEmailAddress, EmailDefaultUnsubscribeURL)
    End Sub
'End EmailDefault Class_Initialize Event

'EmailDefault Initialize Method @33-A5E37F03
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("sesSiteID") = Session("SiteID")
        End With
    End Sub
'End EmailDefault Initialize Method

'EmailDefault Class_Terminate Event @33-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End EmailDefault Class_Terminate Event

'EmailDefault Validate Method @33-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End EmailDefault Validate Method

'EmailDefault Operation Method @33-AABF1E92
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = IIf(EditMode, "Button_Update", "Button_Cancel")
            If Button_Update.Pressed Then
                PressedButton = "Button_Update"
            ElseIf Button_Cancel.Pressed Then
                PressedButton = "Button_Cancel"
            End If
        End If
        Redirect = "AdminEmailSettings.asp?" & CCGetQueryString("QueryString", Array("ccsForm", "Button_Update.x", "Button_Update.y", "Button_Update", "Button_Cancel.x", "Button_Cancel.y", "Button_Cancel"))
        If PressedButton = "Button_Cancel" Then
            If NOT Button_Cancel.OnClick Then
                Redirect = ""
            End If
        ElseIf Validate() Then
            If PressedButton = "Button_Update" Then
                If NOT Button_Update.OnClick() OR NOT UpdateRow() Then
                    Redirect = ""
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End EmailDefault Operation Method

'EmailDefault UpdateRow Method @33-CDB4D9CA
    Function UpdateRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUpdate", Me)
        If NOT UpdateAllowed Then UpdateRow = False : Exit Function
        DataSource.EmailDefaultSiteURL.Value = EmailDefaultSiteURL.Value
        DataSource.EmailDefaultSMTPServer.Value = EmailDefaultSMTPServer.Value
        DataSource.EmailDefaultSMTPPort.Value = EmailDefaultSMTPPort.Value
        DataSource.EmailDefaultSMTPAuthentication.Value = EmailDefaultSMTPAuthentication.Value
        DataSource.EmailDefaultSMTPUserName.Value = EmailDefaultSMTPUserName.Value
        DataSource.EmailDefaultSMTPPassword.Value = EmailDefaultSMTPPassword.Value
        DataSource.EmailDefaultFromAddress.Value = EmailDefaultFromAddress.Value
        DataSource.EmailDefaultFromName.Value = EmailDefaultFromName.Value
        DataSource.EmailDefaultReplyToAddress.Value = EmailDefaultReplyToAddress.Value
        DataSource.EmailDefaultReplyToName.Value = EmailDefaultReplyToName.Value
        DataSource.EmailDefaultBounceBackEmail.Value = EmailDefaultBounceBackEmail.Value
        DataSource.EmailDefaultBounceBackEmailPassword.Value = EmailDefaultBounceBackEmailPassword.Value
        DataSource.EmailDefaultPOPServer.Value = EmailDefaultPOPServer.Value
        DataSource.EmailDefaultPOPPort.Value = EmailDefaultPOPPort.Value
        DataSource.EmailDefaultDNS.Value = EmailDefaultDNS.Value
        DataSource.EmailDefaultDNSHelo.Value = EmailDefaultDNSHelo.Value
        DataSource.EmailDefaultAdminEmailAddress.Value = EmailDefaultAdminEmailAddress.Value
        DataSource.EmailDefaultUnsubscribeURL.Value = EmailDefaultUnsubscribeURL.Value
        DataSource.Update(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterUpdate", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        UpdateRow = (Errors.Count = 0)
    End Function
'End EmailDefault UpdateRow Method

'EmailDefault Show Method @33-DB54E53F
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "EmailDefault" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString
                .Parse False
            End With
        End If
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(EmailDefaultSiteURL,  EmailDefaultSMTPServer,  EmailDefaultSMTPPort,  EmailDefaultSMTPAuthentication,  EmailDefaultSMTPUserName,  EmailDefaultSMTPPassword,  EmailDefaultFromAddress, _
                 EmailDefaultFromName,  EmailDefaultReplyToAddress,  EmailDefaultReplyToName,  EmailDefaultBounceBackEmail,  EmailDefaultBounceBackEmailPassword,  EmailDefaultPOPServer,  EmailDefaultPOPPort,  EmailDefaultDNS, _
                 EmailDefaultDNSHelo,  EmailDefaultAdminEmailAddress,  EmailDefaultUnsubscribeURL,  Button_Update,  Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        EmailDefaultSiteURL.Value = Recordset.Fields("EmailDefaultSiteURL")
                        EmailDefaultSMTPServer.Value = Recordset.Fields("EmailDefaultSMTPServer")
                        EmailDefaultSMTPPort.Value = Recordset.Fields("EmailDefaultSMTPPort")
                        EmailDefaultSMTPAuthentication.Value = Recordset.Fields("EmailDefaultSMTPAuthentication")
                        EmailDefaultSMTPUserName.Value = Recordset.Fields("EmailDefaultSMTPUserName")
                        EmailDefaultSMTPPassword.Value = Recordset.Fields("EmailDefaultSMTPPassword")
                        EmailDefaultFromAddress.Value = Recordset.Fields("EmailDefaultFromAddress")
                        EmailDefaultFromName.Value = Recordset.Fields("EmailDefaultFromName")
                        EmailDefaultReplyToAddress.Value = Recordset.Fields("EmailDefaultReplyToAddress")
                        EmailDefaultReplyToName.Value = Recordset.Fields("EmailDefaultReplyToName")
                        EmailDefaultBounceBackEmail.Value = Recordset.Fields("EmailDefaultBounceBackEmail")
                        EmailDefaultBounceBackEmailPassword.Value = Recordset.Fields("EmailDefaultBounceBackEmailPassword")
                        EmailDefaultPOPServer.Value = Recordset.Fields("EmailDefaultPOPServer")
                        EmailDefaultPOPPort.Value = Recordset.Fields("EmailDefaultPOPPort")
                        EmailDefaultDNS.Value = Recordset.Fields("EmailDefaultDNS")
                        EmailDefaultDNSHelo.Value = Recordset.Fields("EmailDefaultDNSHelo")
                        EmailDefaultAdminEmailAddress.Value = Recordset.Fields("EmailDefaultAdminEmailAddress")
                        EmailDefaultUnsubscribeURL.Value = Recordset.Fields("EmailDefaultUnsubscribeURL")
                    End If
                Else
                    EditMode = False
                End If
            End If
        End If
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors EmailDefaultSiteURL.Errors
            Errors.AddErrors EmailDefaultSMTPServer.Errors
            Errors.AddErrors EmailDefaultSMTPPort.Errors
            Errors.AddErrors EmailDefaultSMTPAuthentication.Errors
            Errors.AddErrors EmailDefaultSMTPUserName.Errors
            Errors.AddErrors EmailDefaultSMTPPassword.Errors
            Errors.AddErrors EmailDefaultFromAddress.Errors
            Errors.AddErrors EmailDefaultFromName.Errors
            Errors.AddErrors EmailDefaultReplyToAddress.Errors
            Errors.AddErrors EmailDefaultReplyToName.Errors
            Errors.AddErrors EmailDefaultBounceBackEmail.Errors
            Errors.AddErrors EmailDefaultBounceBackEmailPassword.Errors
            Errors.AddErrors EmailDefaultPOPServer.Errors
            Errors.AddErrors EmailDefaultPOPPort.Errors
            Errors.AddErrors EmailDefaultDNS.Errors
            Errors.AddErrors EmailDefaultDNSHelo.Errors
            Errors.AddErrors EmailDefaultAdminEmailAddress.Errors
            Errors.AddErrors EmailDefaultUnsubscribeURL.Errors
            Errors.AddErrors DataSource.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction
        Button_Update.Visible = EditMode AND UpdateAllowed

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End EmailDefault Show Method

End Class 'End EmailDefault Class @33-A61BA892

Class clsEmailDefaultDataSource 'EmailDefaultDataSource Class @33-34EA58BA

'DataSource Variables @33-322A869B
    Public Errors, Connection, Parameters, CCSEvents

    Public Recordset
    Public SQL, CountSQL, Order, Where, Orders, StaticOrder
    Public PageSize
    Public PageCount
    Public AbsolutePage
    Public Fields
    Dim WhereParameters
    Public AllParamsSet
    Public CmdExecution

    Private CurrentOperation
    Private CCSEventResult

    ' Datasource fields
    Public EmailDefaultSiteURL
    Public EmailDefaultSMTPServer
    Public EmailDefaultSMTPPort
    Public EmailDefaultSMTPAuthentication
    Public EmailDefaultSMTPUserName
    Public EmailDefaultSMTPPassword
    Public EmailDefaultFromAddress
    Public EmailDefaultFromName
    Public EmailDefaultReplyToAddress
    Public EmailDefaultReplyToName
    Public EmailDefaultBounceBackEmail
    Public EmailDefaultBounceBackEmailPassword
    Public EmailDefaultPOPServer
    Public EmailDefaultPOPPort
    Public EmailDefaultDNS
    Public EmailDefaultDNSHelo
    Public EmailDefaultAdminEmailAddress
    Public EmailDefaultUnsubscribeURL
'End DataSource Variables

'DataSource Class_Initialize Event @33-73BBDB1C
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set EmailDefaultSiteURL = CCCreateField("EmailDefaultSiteURL", "EmailDefaultSiteURL", ccsText, Empty, Recordset)
        Set EmailDefaultSMTPServer = CCCreateField("EmailDefaultSMTPServer", "EmailDefaultSMTPServer", ccsText, Empty, Recordset)
        Set EmailDefaultSMTPPort = CCCreateField("EmailDefaultSMTPPort", "EmailDefaultSMTPPort", ccsInteger, Empty, Recordset)
        Set EmailDefaultSMTPAuthentication = CCCreateField("EmailDefaultSMTPAuthentication", "EmailDefaultSMTPAuthentication", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set EmailDefaultSMTPUserName = CCCreateField("EmailDefaultSMTPUserName", "EmailDefaultSMTPUserName", ccsText, Empty, Recordset)
        Set EmailDefaultSMTPPassword = CCCreateField("EmailDefaultSMTPPassword", "EmailDefaultSMTPPassword", ccsText, Empty, Recordset)
        Set EmailDefaultFromAddress = CCCreateField("EmailDefaultFromAddress", "EmailDefaultFromAddress", ccsText, Empty, Recordset)
        Set EmailDefaultFromName = CCCreateField("EmailDefaultFromName", "EmailDefaultFromName", ccsText, Empty, Recordset)
        Set EmailDefaultReplyToAddress = CCCreateField("EmailDefaultReplyToAddress", "EmailDefaultReplyToAddress", ccsText, Empty, Recordset)
        Set EmailDefaultReplyToName = CCCreateField("EmailDefaultReplyToName", "EmailDefaultReplyToName", ccsText, Empty, Recordset)
        Set EmailDefaultBounceBackEmail = CCCreateField("EmailDefaultBounceBackEmail", "EmailDefaultBounceBackEmail", ccsText, Empty, Recordset)
        Set EmailDefaultBounceBackEmailPassword = CCCreateField("EmailDefaultBounceBackEmailPassword", "EmailDefaultBounceBackEmailPassword", ccsText, Empty, Recordset)
        Set EmailDefaultPOPServer = CCCreateField("EmailDefaultPOPServer", "EmailDefaultPOPServer", ccsText, Empty, Recordset)
        Set EmailDefaultPOPPort = CCCreateField("EmailDefaultPOPPort", "EmailDefaultPOPPort", ccsInteger, Empty, Recordset)
        Set EmailDefaultDNS = CCCreateField("EmailDefaultDNS", "EmailDefaultDNS", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set EmailDefaultDNSHelo = CCCreateField("EmailDefaultDNSHelo", "EmailDefaultDNSHelo", ccsText, Empty, Recordset)
        Set EmailDefaultAdminEmailAddress = CCCreateField("EmailDefaultAdminEmailAddress", "EmailDefaultAdminEmailAddress", ccsText, Empty, Recordset)
        Set EmailDefaultUnsubscribeURL = CCCreateField("EmailDefaultUnsubscribeURL", "EmailDefaultUnsubscribeURL", ccsText, Empty, Recordset)
        Fields.AddFields Array(EmailDefaultSiteURL,  EmailDefaultSMTPServer,  EmailDefaultSMTPPort,  EmailDefaultSMTPAuthentication,  EmailDefaultSMTPUserName,  EmailDefaultSMTPPassword,  EmailDefaultFromAddress, _
             EmailDefaultFromName,  EmailDefaultReplyToAddress,  EmailDefaultReplyToName,  EmailDefaultBounceBackEmail,  EmailDefaultBounceBackEmailPassword,  EmailDefaultPOPServer,  EmailDefaultPOPPort,  EmailDefaultDNS, _
             EmailDefaultDNSHelo,  EmailDefaultAdminEmailAddress,  EmailDefaultUnsubscribeURL)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1  *  " & vbLf & _
        "FROM EmailDefault {SQL_Where} {SQL_OrderBy}"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @33-461C4BAD
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "sesSiteID", ccsInteger, Empty, Empty, -1, False
            AllParamsSet = .AllParamsSet
            .Criterion(1) = .Operation(opEqual, False, "[EmailDefaultSiteID]", .getParamByID(1))
            .AssembledWhere = .Criterion(1)
            WhereParams = .AssembledWhere
            If Len(Where) > 0 Then 
                If Len(WhereParams) > 0 Then _
                    Where = Where & " AND " & WhereParams
            Else
                If Len(WhereParams) > 0 Then _
                    Where = WhereParams
            End If
        End With
    End Sub
'End BuildTableWhere Method

'Open Method @33-48A2DA7D
    Function Open(Cmd)
        Errors.Clear
        If Connection Is Nothing Then
            Set Open = New clsEmptyDataSource
            Exit Function
        End If
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdOpen
        Cmd.PageSize = PageSize
        Cmd.ActivePage = AbsolutePage
        Cmd.CommandType = dsTable
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildSelect", Me)
        Cmd.SQL = SQL
        BuildTableWhere
        Cmd.Where = Where
        Cmd.OrderBy = Order
        If(Len(StaticOrder)>0) Then
            If Len(Order)>0 Then Cmd.OrderBy = ", "+Cmd.OrderBy
            Cmd.OrderBy = StaticOrder + Cmd.OrderBy
        End If
        Cmd.Options("TOP") = True
        If Not AllParamsSet Then
            Set Open = New clsEmptyDataSource
            Exit Function
        End If
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteSelect", Me)
        If Errors.Count = 0 And CCSEventResult Then _
            Set Recordset = Cmd.Exec(Errors)
        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteSelect", Me)
        Set Recordset.FieldsCollection = Fields
        Set Open = Recordset
    End Function
'End Open Method

'DataSource Class_Terminate Event @33-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

'Update Method @33-0B494C6F
    Sub Update(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildUpdate", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        BuildTableWhere
        If NOT AllParamsSet Then
            Errors.AddError(CCSLocales.GetText("CCS_CustomOperationError_MissingParameters", Empty))
        End If
        Cmd.SQL = "UPDATE [EmailDefault] SET " & _
            "[EmailDefaultSiteURL]=" & Connection.ToSQL(EmailDefaultSiteURL, EmailDefaultSiteURL.DataType) & ", " & _
            "[EmailDefaultSMTPServer]=" & Connection.ToSQL(EmailDefaultSMTPServer, EmailDefaultSMTPServer.DataType) & ", " & _
            "[EmailDefaultSMTPPort]=" & Connection.ToSQL(EmailDefaultSMTPPort, EmailDefaultSMTPPort.DataType) & ", " & _
            "[EmailDefaultSMTPAuthentication]=" & Connection.ToSQL(EmailDefaultSMTPAuthentication, EmailDefaultSMTPAuthentication.DataType) & ", " & _
            "[EmailDefaultSMTPUserName]=" & Connection.ToSQL(EmailDefaultSMTPUserName, EmailDefaultSMTPUserName.DataType) & ", " & _
            "[EmailDefaultSMTPPassword]=" & Connection.ToSQL(EmailDefaultSMTPPassword, EmailDefaultSMTPPassword.DataType) & ", " & _
            "[EmailDefaultFromAddress]=" & Connection.ToSQL(EmailDefaultFromAddress, EmailDefaultFromAddress.DataType) & ", " & _
            "[EmailDefaultFromName]=" & Connection.ToSQL(EmailDefaultFromName, EmailDefaultFromName.DataType) & ", " & _
            "[EmailDefaultReplyToAddress]=" & Connection.ToSQL(EmailDefaultReplyToAddress, EmailDefaultReplyToAddress.DataType) & ", " & _
            "[EmailDefaultReplyToName]=" & Connection.ToSQL(EmailDefaultReplyToName, EmailDefaultReplyToName.DataType) & ", " & _
            "[EmailDefaultBounceBackEmail]=" & Connection.ToSQL(EmailDefaultBounceBackEmail, EmailDefaultBounceBackEmail.DataType) & ", " & _
            "[EmailDefaultBounceBackEmailPassword]=" & Connection.ToSQL(EmailDefaultBounceBackEmailPassword, EmailDefaultBounceBackEmailPassword.DataType) & ", " & _
            "[EmailDefaultPOPServer]=" & Connection.ToSQL(EmailDefaultPOPServer, EmailDefaultPOPServer.DataType) & ", " & _
            "[EmailDefaultPOPPort]=" & Connection.ToSQL(EmailDefaultPOPPort, EmailDefaultPOPPort.DataType) & ", " & _
            "[EmailDefaultDNS]=" & Connection.ToSQL(EmailDefaultDNS, EmailDefaultDNS.DataType) & ", " & _
            "[EmailDefaultDNSHelo]=" & Connection.ToSQL(EmailDefaultDNSHelo, EmailDefaultDNSHelo.DataType) & ", " & _
            "[EmailDefaultAdminEmailAddress]=" & Connection.ToSQL(EmailDefaultAdminEmailAddress, EmailDefaultAdminEmailAddress.DataType) & ", " & _
            "[EmailDefaultUnsubscribeURL]=" & Connection.ToSQL(EmailDefaultUnsubscribeURL, EmailDefaultUnsubscribeURL.DataType) & _
            IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

End Class 'End EmailDefaultDataSource Class @33-A61BA892


%>
