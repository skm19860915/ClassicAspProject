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

'Initialize Page @1-428E2D25
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
Dim UserTable
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminUserTableEdit.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminUserTableEdit.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-6FCE7A5C
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set UserTable = new clsRecordUserTable
Menu.Value = DHTMLMenu

UserTable.Initialize DBSystem

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-1D056FCE
UserTable.Operation
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

'Show Page @1-069B40E5
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, UserTable))
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

'UnloadPage Sub @1-FC5080A1
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set UserTable = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordUserTable 'UserTable Class @56-31CA43BA

'UserTable Variables @56-4DAFAA16

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
    Dim UserTableLogin
    Dim UserTablePassword
    Dim UserTableEmail
    Dim UserTableUserGroupID
    Dim SiteList
    Dim UserTableSiteList
    Dim UserTableActive
    Dim Button_Insert
    Dim Button_Update
    Dim Button_Cancel
'End UserTable Variables

'UserTable Class_Initialize Event @56-8DFDD5F4
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsUserTableDataSource
        Set Command = New clsCommand
        InsertAllowed = True
        UpdateAllowed = True
        DeleteAllowed = False
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "UserTable")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "UserTable"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set UserTableLogin = CCCreateControl(ccsTextBox, "UserTableLogin", "Login", ccsText, Empty, CCGetRequestParam("UserTableLogin", Method))
        Set UserTablePassword = CCCreateControl(ccsTextBox, "UserTablePassword", "Password", ccsText, Empty, CCGetRequestParam("UserTablePassword", Method))
        Set UserTableEmail = CCCreateControl(ccsTextBox, "UserTableEmail", "Email", ccsText, Empty, CCGetRequestParam("UserTableEmail", Method))
        Set UserTableUserGroupID = CCCreateList(ccsListBox, "UserTableUserGroupID", "User Group ID", ccsInteger, CCGetRequestParam("UserTableUserGroupID", Method), Empty)
        Set UserTableUserGroupID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM UserGroup {SQL_Where} {SQL_OrderBy}", "", ""))
        Set SiteList = CCCreateList(ccsListBox, "SiteList", "Site List", ccsText, CCGetRequestMultipleParam("SiteList", Method), Empty)
        SiteList.IsMultiple = True
        Set SiteList.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM Site {SQL_Where} {SQL_OrderBy}", "", ""))
        Set UserTableSiteList = CCCreateControl(ccsHidden, "UserTableSiteList", Empty, ccsText, Empty, CCGetRequestParam("UserTableSiteList", Method))
        Set UserTableActive = CCCreateControl(ccsCheckBox, "UserTableActive", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("UserTableActive", Method))
        Set Button_Insert = CCCreateButton("Button_Insert", Method)
        Set Button_Update = CCCreateButton("Button_Update", Method)
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(UserTableLogin, UserTablePassword, UserTableEmail, UserTableUserGroupID, SiteList, UserTableSiteList, UserTableActive)
    End Sub
'End UserTable Class_Initialize Event

'UserTable Initialize Method @56-2AD50A9A
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("urlUserTableID") = CCGetRequestParam("UserTableID", ccsGET)
        End With
    End Sub
'End UserTable Initialize Method

'UserTable Class_Terminate Event @56-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End UserTable Class_Terminate Event

'UserTable Validate Method @56-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End UserTable Validate Method

'UserTable Operation Method @56-EE316774
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = IIf(EditMode, "Button_Update", "Button_Insert")
            If Button_Insert.Pressed Then
                PressedButton = "Button_Insert"
            ElseIf Button_Update.Pressed Then
                PressedButton = "Button_Update"
            ElseIf Button_Cancel.Pressed Then
                PressedButton = "Button_Cancel"
            End If
        End If
        Redirect = "AdminUserTableList.asp"
        If PressedButton = "Button_Cancel" Then
            If NOT Button_Cancel.OnClick Then
                Redirect = ""
            End If
        ElseIf Validate() Then
            If PressedButton = "Button_Insert" Then
                If NOT Button_Insert.OnClick() OR NOT InsertRow() Then
                    Redirect = ""
                End If
            ElseIf PressedButton = "Button_Update" Then
                If NOT Button_Update.OnClick() OR NOT UpdateRow() Then
                    Redirect = ""
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End UserTable Operation Method

'UserTable InsertRow Method @56-4A1DE044
    Function InsertRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeInsert", Me)
        If NOT InsertAllowed Then InsertRow = False : Exit Function
        DataSource.UserTableLogin.Value = UserTableLogin.Value
        DataSource.UserTablePassword.Value = UserTablePassword.Value
        DataSource.UserTableEmail.Value = UserTableEmail.Value
        DataSource.UserTableUserGroupID.Value = UserTableUserGroupID.Value
        DataSource.UserTableSiteList.Value = UserTableSiteList.Value
        DataSource.UserTableActive.Value = UserTableActive.Value
        DataSource.Insert(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInsert", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        InsertRow = (Errors.Count = 0)
    End Function
'End UserTable InsertRow Method

'UserTable UpdateRow Method @56-59D0C43C
    Function UpdateRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUpdate", Me)
        If NOT UpdateAllowed Then UpdateRow = False : Exit Function
        DataSource.UserTableLogin.Value = UserTableLogin.Value
        DataSource.UserTablePassword.Value = UserTablePassword.Value
        DataSource.UserTableEmail.Value = UserTableEmail.Value
        DataSource.UserTableUserGroupID.Value = UserTableUserGroupID.Value
        DataSource.UserTableSiteList.Value = UserTableSiteList.Value
        DataSource.UserTableActive.Value = UserTableActive.Value
        DataSource.Update(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterUpdate", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        UpdateRow = (Errors.Count = 0)
    End Function
'End UserTable UpdateRow Method

'UserTable Show Method @56-281D1CEA
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "UserTable" & IIf(EditMode, ":Edit", ""))
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
            Array(UserTableLogin,  UserTablePassword,  UserTableEmail,  UserTableUserGroupID,  SiteList,  UserTableSiteList,  UserTableActive, _
                 Button_Insert,  Button_Update,  Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        UserTableLogin.Value = Recordset.Fields("UserTableLogin")
                        UserTablePassword.Value = Recordset.Fields("UserTablePassword")
                        UserTableEmail.Value = Recordset.Fields("UserTableEmail")
                        UserTableUserGroupID.Value = Recordset.Fields("UserTableUserGroupID")
                        
                        UserTableSiteList.Value = Recordset.Fields("UserTableSiteList")
                        UserTableActive.Value = Recordset.Fields("UserTableActive")
                    End If
                Else
                    EditMode = False
                End If
            End If
        End If
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors UserTableLogin.Errors
            Errors.AddErrors UserTablePassword.Errors
            Errors.AddErrors UserTableEmail.Errors
            Errors.AddErrors UserTableUserGroupID.Errors
            Errors.AddErrors SiteList.Errors
            Errors.AddErrors UserTableSiteList.Errors
            Errors.AddErrors UserTableActive.Errors
            Errors.AddErrors DataSource.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction
        Button_Insert.Visible = NOT EditMode AND InsertAllowed
        Button_Update.Visible = EditMode AND UpdateAllowed

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End UserTable Show Method

End Class 'End UserTable Class @56-A61BA892

Class clsUserTableDataSource 'UserTableDataSource Class @56-8AC67EAD

'DataSource Variables @56-B84D2F06
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
    Public UserTableLogin
    Public UserTablePassword
    Public UserTableEmail
    Public UserTableUserGroupID
    Public UserTableSiteList
    Public UserTableActive
'End DataSource Variables

'DataSource Class_Initialize Event @56-7D95E95F
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set UserTableLogin = CCCreateField("UserTableLogin", "UserTableLogin", ccsText, Empty, Recordset)
        Set UserTablePassword = CCCreateField("UserTablePassword", "UserTablePassword", ccsText, Empty, Recordset)
        Set UserTableEmail = CCCreateField("UserTableEmail", "UserTableEmail", ccsText, Empty, Recordset)
        Set UserTableUserGroupID = CCCreateField("UserTableUserGroupID", "UserTableUserGroupID", ccsInteger, Empty, Recordset)
        Set UserTableSiteList = CCCreateField("UserTableSiteList", "UserTableSiteList", ccsText, Empty, Recordset)
        Set UserTableActive = CCCreateField("UserTableActive", "UserTableActive", ccsBoolean, Array(1, 0, Empty), Recordset)
        Fields.AddFields Array(UserTableLogin, UserTablePassword, UserTableEmail, UserTableUserGroupID, UserTableSiteList, UserTableActive)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1  *  " & vbLf & _
        "FROM UserTable {SQL_Where} {SQL_OrderBy}"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @56-C1856182
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urlUserTableID", ccsInteger, Empty, Empty, Empty, False
            AllParamsSet = .AllParamsSet
            .Criterion(1) = .Operation(opEqual, False, "[UserTableID]", .getParamByID(1))
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

'Open Method @56-48A2DA7D
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

'DataSource Class_Terminate Event @56-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

'Update Method @56-012B46E2
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
        Cmd.SQL = "UPDATE [UserTable] SET " & _
            "[UserTableLogin]=" & Connection.ToSQL(UserTableLogin, UserTableLogin.DataType) & ", " & _
            "[UserTablePassword]=" & Connection.ToSQL(UserTablePassword, UserTablePassword.DataType) & ", " & _
            "[UserTableEmail]=" & Connection.ToSQL(UserTableEmail, UserTableEmail.DataType) & ", " & _
            "[UserTableUserGroupID]=" & Connection.ToSQL(UserTableUserGroupID, UserTableUserGroupID.DataType) & ", " & _
            "[UserTableSiteList]=" & Connection.ToSQL(UserTableSiteList, UserTableSiteList.DataType) & ", " & _
            "[UserTableActive]=" & Connection.ToSQL(UserTableActive, UserTableActive.DataType) & _
            IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

'Insert Method @56-B18922FE
    Sub Insert(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildInsert", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Cmd.SQL = "INSERT INTO [UserTable] (" & _
            "[UserTableLogin], " & _
            "[UserTablePassword], " & _
            "[UserTableEmail], " & _
            "[UserTableUserGroupID], " & _
            "[UserTableSiteList], " & _
            "[UserTableActive]" & _
        ") VALUES (" & _
            Connection.ToSQL(UserTableLogin, UserTableLogin.DataType) & ", " & _
            Connection.ToSQL(UserTablePassword, UserTablePassword.DataType) & ", " & _
            Connection.ToSQL(UserTableEmail, UserTableEmail.DataType) & ", " & _
            Connection.ToSQL(UserTableUserGroupID, UserTableUserGroupID.DataType) & ", " & _
            Connection.ToSQL(UserTableSiteList, UserTableSiteList.DataType) & ", " & _
            Connection.ToSQL(UserTableActive, UserTableActive.DataType) & _
        ")"
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteInsert", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteInsert", Me)
        End If
    End Sub
'End Insert Method

End Class 'End UserTableDataSource Class @56-A61BA892


%>
