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

'Initialize Page @1-2897F036
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
Dim UserTableSearch
Dim UserTable
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminUserTableList.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminUserTableList.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-1B327F1F
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set UserTableSearch = new clsRecordUserTableSearch
Set UserTable = New clsGridUserTable
Menu.Value = DHTMLMenu

UserTable.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminUserTableList_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-B4703A1D
UserTableSearch.Operation
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

'Show Page @1-6EA25F01
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, UserTableSearch, UserTable))
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

'UnloadPage Sub @1-6BDE3B8B
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set UserTableSearch = Nothing
    Set UserTable = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordUserTableSearch 'UserTableSearch Class @34-2B074DDE

'UserTableSearch Variables @34-6F407D4E

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
    Dim searchString
    Dim UserTablePageSize
    Dim Button_DoSearch
'End UserTableSearch Variables

'UserTableSearch Class_Initialize Event @34-0CE62B45
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        InsertAllowed = False
        UpdateAllowed = False
        DeleteAllowed = False
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "UserTableSearch")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "UserTableSearch"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set searchString = CCCreateControl(ccsTextBox, "searchString", Empty, ccsText, Empty, CCGetRequestParam("searchString", Method))
        Set UserTablePageSize = CCCreateList(ccsListBox, "UserTablePageSize", Empty, ccsText, CCGetRequestParam("UserTablePageSize", Method), Empty)
        Set UserTablePageSize.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("", "5", "10", "25", "100"), _
            Array("Select Value", "5", "10", "25", "100")))
        Set Button_DoSearch = CCCreateButton("Button_DoSearch", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(searchString, UserTablePageSize)
    End Sub
'End UserTableSearch Class_Initialize Event

'UserTableSearch Class_Terminate Event @34-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End UserTableSearch Class_Terminate Event

'UserTableSearch Validate Method @34-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End UserTableSearch Validate Method

'UserTableSearch Operation Method @34-66557847
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoSearch"
            If Button_DoSearch.Pressed Then
                PressedButton = "Button_DoSearch"
            End If
        End If
        Redirect = "AdminUserTableList.asp"
        If Validate() Then
            If PressedButton = "Button_DoSearch" Then
                If NOT Button_DoSearch.OnClick() Then
                    Redirect = ""
                Else
                    Redirect = "AdminUserTableList.asp?" & CCGetQueryString("Form", Array(PressedButton, "ccsForm", "Button_DoSearch.x", "Button_DoSearch.y", "Button_DoSearch"))
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End UserTableSearch Operation Method

'UserTableSearch Show Method @34-8895EA82
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "UserTableSearch" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(searchString, UserTablePageSize, Button_DoSearch))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors searchString.Errors
            Errors.AddErrors UserTablePageSize.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End UserTableSearch Show Method

End Class 'End UserTableSearch Class @34-A61BA892

Class clsGridUserTable 'UserTable Class @33-9485CC88

'UserTable Variables @33-41934FD4

    ' Private variables
    Private VarPageSize
    ' Public variables
    Public ComponentName, CCSEvents
    Public Visible, Errors
    Public DataSource
    Public PageNumber
    Public Command
    Public TemplateBlock
    Public IsDSEmpty
    Public ForceIteration
    Private ShownRecords
    Public ActiveSorter, SortingDirection
    Public Recordset

    Private CCSEventResult

    ' Grid Controls
    Public StaticControls, RowControls
    Public AltRowControls
    Public RenderAltRow
    Dim Link1
    Dim UserTable_TotalRecords
    Dim Sorter_UserLogin
    Dim Sorter_UserGroupName
    Dim UserLogin
    Dim UserGroupName
    Dim UserTableActive
    Dim Alt_UserLogin
    Dim Alt_UserGroupName
    Dim Alt_UserTableActive
    Dim Navigator
'End UserTable Variables

'UserTable Class_Initialize Event @33-E063F252
    Private Sub Class_Initialize()
        ComponentName = "UserTable"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsUserTableDataSource
        Set Command = New clsCommand
        PageSize = CCGetParam(ComponentName & "PageSize", Empty)
        If IsNumeric(PageSize) And Len(PageSize) > 0 Then
            If PageSize <= 0 Then Errors.AddError(CCSLocales.GetText("GridPageSizeError", Empty))
            If PageSize > 100 Then PageSize = 100
        End If
        If NOT IsNumeric(PageSize) OR IsEmpty(PageSize) Then _
            PageSize = 25 _
        Else _
            PageSize = CInt(PageSize)
        PageNumber = CCGetParam(ComponentName & "Page", 1)
        If Not IsNumeric(PageNumber) And Len(PageNumber) > 0 Then
            Errors.AddError(CCSLocales.GetText("GridPageNumberError", Empty))
            PageNumber = 1
        ElseIf Len(PageNumber) > 0 Then
            If PageNumber > 0 Then
                PageNumber = CInt(PageNumber)
            Else
                Errors.AddError(CCSLocales.GetText("GridPageNumberError", Empty))
                PageNumber = 1
            End If
        Else
            PageNumber = 1
        End If
        ActiveSorter = CCGetParam("UserTableOrder", Empty)
        SortingDirection = CCGetParam("UserTableDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Link1 = CCCreateControl(ccsLink, "Link1", Empty, ccsText, Empty, CCGetRequestParam("Link1", ccsGet))
        Set UserTable_TotalRecords = CCCreateControl(ccsLabel, "UserTable_TotalRecords", Empty, ccsText, Empty, CCGetRequestParam("UserTable_TotalRecords", ccsGet))
        Set Sorter_UserLogin = CCCreateSorter("Sorter_UserLogin", Me, FileName)
        Set Sorter_UserGroupName = CCCreateSorter("Sorter_UserGroupName", Me, FileName)
        Set UserLogin = CCCreateControl(ccsLink, "UserLogin", Empty, ccsText, Empty, CCGetRequestParam("UserLogin", ccsGet))
        Set UserGroupName = CCCreateControl(ccsLabel, "UserGroupName", Empty, ccsText, Empty, CCGetRequestParam("UserGroupName", ccsGet))
        Set UserTableActive = CCCreateControl(ccsLabel, "UserTableActive", Empty, ccsBoolean, Array("Yes", "No", Empty), CCGetRequestParam("UserTableActive", ccsGet))
        Set Alt_UserLogin = CCCreateControl(ccsLink, "Alt_UserLogin", Empty, ccsText, Empty, CCGetRequestParam("Alt_UserLogin", ccsGet))
        Set Alt_UserGroupName = CCCreateControl(ccsLabel, "Alt_UserGroupName", Empty, ccsText, Empty, CCGetRequestParam("Alt_UserGroupName", ccsGet))
        Set Alt_UserTableActive = CCCreateControl(ccsLabel, "Alt_UserTableActive", Empty, ccsBoolean, Array("Yes", "No", Empty), CCGetRequestParam("Alt_UserTableActive", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
    IsDSEmpty = True
    End Sub
'End UserTable Class_Initialize Event

'UserTable Initialize Method @33-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End UserTable Initialize Method

'UserTable Class_Terminate Event @33-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End UserTable Class_Terminate Event

'UserTable Show Method @33-333E1E4D
    Sub Show(Tpl)
        Dim HasNext
        If NOT Visible Then Exit Sub

        Dim RowBlock, AltRowBlock

        With DataSource
            .Parameters("urlsearchString") = CCGetRequestParam("searchString", ccsGET)
        End With

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        IsDSEmpty = Recordset.EOF

        Set TemplateBlock = Tpl.Block("Grid " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        Set RowBlock = TemplateBlock.Block("Row")
        Set AltRowBlock = TemplateBlock.Block("AltRow")
        Set StaticControls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(Link1, UserTable_TotalRecords, Sorter_UserLogin, Sorter_UserGroupName, Navigator))
            
            Link1.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
            Link1.Page = "AdminUserTableEdit.asp"
            
            Navigator.SetDataSource Recordset
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(UserLogin, UserGroupName, UserTableActive))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_UserLogin, Alt_UserGroupName, Alt_UserTableActive))

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If NOT Visible Then Exit Sub

        RowControls.PreserveControlsVisible
        AltRowControls.PreserveControlsVisible
        Errors.AddErrors DataSource.Errors
        If Errors.Count > 0 Then
            TemplateBlock.HTML = CCFormatError("Grid " & ComponentName, Errors)
        Else

            ' Show NoRecords block if no records are found
            If Recordset.EOF Then
                TemplateBlock.Block("NoRecords").Parse ccsParseOverwrite
            End If
            ForceIteration= HasNextRow()
            Do While ForceIteration
                HasNext = HasNextRow()
                If RenderAltRow Then
                    If HasNext Then
                        Alt_UserLogin.Value = Recordset.Fields("Alt_UserLogin")
                        Alt_UserLogin.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        Alt_UserLogin.Parameters = CCAddParam(Alt_UserLogin.Parameters, "UserTableID", Recordset.Fields("Alt_UserLogin_param1"))
                        Alt_UserLogin.Page = "AdminUserTableEdit.asp"
                        Alt_UserGroupName.Value = Recordset.Fields("Alt_UserGroupName")
                        Alt_UserTableActive.Value = Recordset.Fields("Alt_UserTableActive")
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    AltRowControls.Show
                Else
                    If HasNext Then
                        UserLogin.Value = Recordset.Fields("UserLogin")
                        UserLogin.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        UserLogin.Parameters = CCAddParam(UserLogin.Parameters, "UserTableID", Recordset.Fields("UserLogin_param1"))
                        UserLogin.Page = "AdminUserTableEdit.asp"
                        UserGroupName.Value = Recordset.Fields("UserGroupName")
                        UserTableActive.Value = Recordset.Fields("UserTableActive")
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    RowControls.Show
                End If
                RenderAltRow = NOT RenderAltRow
                If HasNext Then Recordset.MoveNext
                ShownRecords = ShownRecords + 1
            Loop
            StaticControls.Show
        End If

    End Sub
'End UserTable Show Method

'UserTable PageSize Property Let @33-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End UserTable PageSize Property Let

'UserTable PageSize Property Get @33-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End UserTable PageSize Property Get

'UserTable RowNumber Property Get @33-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End UserTable RowNumber Property Get

'UserTable HasNextRow Function @33-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End UserTable HasNextRow Function

End Class 'End UserTable Class @33-A61BA892

Class clsUserTableDataSource 'UserTableDataSource Class @33-8AC67EAD

'DataSource Variables @33-D2079DB4
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
    Public UserLogin
    Public UserLogin_param1
    Public UserGroupName
    Public UserTableActive
    Public Alt_UserLogin
    Public Alt_UserLogin_param1
    Public Alt_UserGroupName
    Public Alt_UserTableActive
'End DataSource Variables

'DataSource Class_Initialize Event @33-139F6EBD
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set UserLogin = CCCreateField("UserLogin", "UserTableLogin", ccsText, Empty, Recordset)
        Set UserLogin_param1 = CCCreateField("UserLogin_param1", "UserTableID", ccsText, Empty, Recordset)
        Set UserGroupName = CCCreateField("UserGroupName", "UserGroupName", ccsText, Empty, Recordset)
        Set UserTableActive = CCCreateField("UserTableActive", "UserTableActive", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set Alt_UserLogin = CCCreateField("Alt_UserLogin", "UserTableLogin", ccsText, Empty, Recordset)
        Set Alt_UserLogin_param1 = CCCreateField("Alt_UserLogin_param1", "UserTableID", ccsText, Empty, Recordset)
        Set Alt_UserGroupName = CCCreateField("Alt_UserGroupName", "UserGroupName", ccsText, Empty, Recordset)
        Set Alt_UserTableActive = CCCreateField("Alt_UserTableActive", "UserTableActive", ccsBoolean, Array(1, 0, Empty), Recordset)
        Fields.AddFields Array(UserLogin, UserLogin_param1, UserGroupName, UserTableActive, Alt_UserLogin, Alt_UserLogin_param1, Alt_UserGroupName, Alt_UserTableActive)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_UserLogin", "UserLogin", ""), _
            Array("Sorter_UserGroupName", "UserGroupName", ""))

        SQL = "SELECT TOP {SqlParam_endRecord}  *  " & vbLf & _
        "FROM UserTable INNER JOIN UserGroup ON " & vbLf & _
        "UserTable.UserTableUserGroupID = UserGroup.UserGroupID {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM UserTable INNER JOIN UserGroup ON " & vbLf & _
        "UserTable.UserTableUserGroupID = UserGroup.UserGroupID"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'SetOrder Method @33-68FC9576
    Sub SetOrder(Column, Direction)
        Order = Recordset.GetOrder(Order, Column, Direction, Orders)
    End Sub
'End SetOrder Method

'BuildTableWhere Method @33-9CC0D603
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urlsearchString", ccsText, Empty, Empty, Empty, False
            .AddParameter 2, "urlsearchString", ccsText, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opContains, False, "[UserTable].[UserTableLogin]", .getParamByID(1))
            .Criterion(2) = .Operation(opContains, False, "[UserGroup].[UserGroupName]", .getParamByID(2))
            .AssembledWhere = .opOR(True, .Criterion(1), .Criterion(2))
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

'Open Method @33-40984FC5
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
        Cmd.CountSQL = CountSQL
        BuildTableWhere
        Cmd.Where = Where
        Cmd.OrderBy = Order
        If(Len(StaticOrder)>0) Then
            If Len(Order)>0 Then Cmd.OrderBy = ", "+Cmd.OrderBy
            Cmd.OrderBy = StaticOrder + Cmd.OrderBy
        End If
        Cmd.Options("TOP") = True
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

End Class 'End UserTableDataSource Class @33-A61BA892


%>
