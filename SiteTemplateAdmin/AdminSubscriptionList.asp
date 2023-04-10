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

'Initialize Page @1-F0F7924E
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
Dim SubscriptionSearch
Dim Subscription
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminSubscriptionList.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminSubscriptionList.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Initialize Objects @1-97661EB3
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set SubscriptionSearch = new clsRecordSubscriptionSearch
Set Subscription = New clsGridSubscription
Menu.Value = DHTMLMenu

Subscription.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminSubscriptionList_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-F998E276
SubscriptionSearch.Operation
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

'Show Page @1-9AA9F4F8
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, SubscriptionSearch, Subscription))
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

'UnloadPage Sub @1-79AA6C2E
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set SubscriptionSearch = Nothing
    Set Subscription = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordSubscriptionSearch 'SubscriptionSearch Class @34-2C8856A8

'SubscriptionSearch Variables @34-2B94D881

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
    Dim s_keyword
    Dim SubscriptionPageSize
    Dim Button_DoSearch
'End SubscriptionSearch Variables

'SubscriptionSearch Class_Initialize Event @34-7CA1987F
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
            FormSubmitted = (OperationMode(0) = "SubscriptionSearch")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "SubscriptionSearch"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set s_keyword = CCCreateControl(ccsTextBox, "s_keyword", Empty, ccsText, Empty, CCGetRequestParam("s_keyword", Method))
        Set SubscriptionPageSize = CCCreateList(ccsListBox, "SubscriptionPageSize", Empty, ccsText, CCGetRequestParam("SubscriptionPageSize", Method), Empty)
        Set SubscriptionPageSize.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("", "5", "10", "25", "100"), _
            Array("Select Value", "5", "10", "25", "100")))
        Set Button_DoSearch = CCCreateButton("Button_DoSearch", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(s_keyword, SubscriptionPageSize)
    End Sub
'End SubscriptionSearch Class_Initialize Event

'SubscriptionSearch Class_Terminate Event @34-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End SubscriptionSearch Class_Terminate Event

'SubscriptionSearch Validate Method @34-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End SubscriptionSearch Validate Method

'SubscriptionSearch Operation Method @34-FE3B6FF1
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoSearch"
            If Button_DoSearch.Pressed Then
                PressedButton = "Button_DoSearch"
            End If
        End If
        Redirect = "AdminSubscriptionList.asp"
        If Validate() Then
            If PressedButton = "Button_DoSearch" Then
                If NOT Button_DoSearch.OnClick() Then
                    Redirect = ""
                Else
                    Redirect = "AdminSubscriptionList.asp?" & CCGetQueryString("Form", Array(PressedButton, "ccsForm", "Button_DoSearch.x", "Button_DoSearch.y", "Button_DoSearch"))
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End SubscriptionSearch Operation Method

'SubscriptionSearch Show Method @34-9EF20EDF
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "SubscriptionSearch" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(s_keyword, SubscriptionPageSize, Button_DoSearch))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors s_keyword.Errors
            Errors.AddErrors SubscriptionPageSize.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End SubscriptionSearch Show Method

End Class 'End SubscriptionSearch Class @34-A61BA892

Class clsGridSubscription 'Subscription Class @33-E9F0B1DF

'Subscription Variables @33-8B985772

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
    Dim Subscription_TotalRecords
    Dim Sorter_SubscriptionID
    Dim Sorter_SubscriptionType
    Dim Sorter_SubscriptionName
    Dim Sorter_SubscriptionDateTimeCreated
    Dim SubscriptionID
    Dim SubscriptionType
    Dim SubscriptionName
    Dim SubscriptionDateTimeCreated
    Dim Alt_SubscriptionID
    Dim Alt_SubscriptionType
    Dim Alt_SubscriptionName
    Dim Alt_SubscriptionDateTimeCreated
    Dim Navigator
'End Subscription Variables

'Subscription Class_Initialize Event @33-9DB5612C
    Private Sub Class_Initialize()
        ComponentName = "Subscription"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsSubscriptionDataSource
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
        ActiveSorter = CCGetParam("SubscriptionOrder", Empty)
        SortingDirection = CCGetParam("SubscriptionDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Subscription_TotalRecords = CCCreateControl(ccsLabel, "Subscription_TotalRecords", Empty, ccsText, Empty, CCGetRequestParam("Subscription_TotalRecords", ccsGet))
        Set Sorter_SubscriptionID = CCCreateSorter("Sorter_SubscriptionID", Me, FileName)
        Set Sorter_SubscriptionType = CCCreateSorter("Sorter_SubscriptionType", Me, FileName)
        Set Sorter_SubscriptionName = CCCreateSorter("Sorter_SubscriptionName", Me, FileName)
        Set Sorter_SubscriptionDateTimeCreated = CCCreateSorter("Sorter_SubscriptionDateTimeCreated", Me, FileName)
        Set SubscriptionID = CCCreateControl(ccsLink, "SubscriptionID", Empty, ccsInteger, Empty, CCGetRequestParam("SubscriptionID", ccsGet))
        Set SubscriptionType = CCCreateControl(ccsLabel, "SubscriptionType", Empty, ccsText, Empty, CCGetRequestParam("SubscriptionType", ccsGet))
        Set SubscriptionName = CCCreateControl(ccsLabel, "SubscriptionName", Empty, ccsText, Empty, CCGetRequestParam("SubscriptionName", ccsGet))
        Set SubscriptionDateTimeCreated = CCCreateControl(ccsLabel, "SubscriptionDateTimeCreated", Empty, ccsDate, Array("GeneralDate"), CCGetRequestParam("SubscriptionDateTimeCreated", ccsGet))
        Set Alt_SubscriptionID = CCCreateControl(ccsLink, "Alt_SubscriptionID", Empty, ccsInteger, Empty, CCGetRequestParam("Alt_SubscriptionID", ccsGet))
        Set Alt_SubscriptionType = CCCreateControl(ccsLabel, "Alt_SubscriptionType", Empty, ccsText, Empty, CCGetRequestParam("Alt_SubscriptionType", ccsGet))
        Set Alt_SubscriptionName = CCCreateControl(ccsLabel, "Alt_SubscriptionName", Empty, ccsText, Empty, CCGetRequestParam("Alt_SubscriptionName", ccsGet))
        Set Alt_SubscriptionDateTimeCreated = CCCreateControl(ccsLabel, "Alt_SubscriptionDateTimeCreated", Empty, ccsDate, Array("GeneralDate"), CCGetRequestParam("Alt_SubscriptionDateTimeCreated", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
    IsDSEmpty = True
    End Sub
'End Subscription Class_Initialize Event

'Subscription Initialize Method @33-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End Subscription Initialize Method

'Subscription Class_Terminate Event @33-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End Subscription Class_Terminate Event

'Subscription Show Method @33-CE1885B0
    Sub Show(Tpl)
        Dim HasNext
        If NOT Visible Then Exit Sub

        Dim RowBlock, AltRowBlock

        With DataSource
            .Parameters("sesSiteID") = Session("SiteID")
            .Parameters("urls_keyword") = CCGetRequestParam("s_keyword", ccsGET)
        End With

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        IsDSEmpty = Recordset.EOF

        Set TemplateBlock = Tpl.Block("Grid " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        Set RowBlock = TemplateBlock.Block("Row")
        Set AltRowBlock = TemplateBlock.Block("AltRow")
        Set StaticControls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(Subscription_TotalRecords, Sorter_SubscriptionID, Sorter_SubscriptionType, Sorter_SubscriptionName, Sorter_SubscriptionDateTimeCreated, Navigator))
            
            Navigator.SetDataSource Recordset
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(SubscriptionID, SubscriptionType, SubscriptionName, SubscriptionDateTimeCreated))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_SubscriptionID, Alt_SubscriptionType, Alt_SubscriptionName, Alt_SubscriptionDateTimeCreated))

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
                        Alt_SubscriptionID.Value = Recordset.Fields("Alt_SubscriptionID")
                        Alt_SubscriptionID.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        Alt_SubscriptionID.Parameters = CCAddParam(Alt_SubscriptionID.Parameters, "SubscriptionID", Recordset.Fields("Alt_SubscriptionID_param1"))
                        Alt_SubscriptionID.Page = "AdminSubscriptionEdit.asp"
                        Alt_SubscriptionType.Value = Recordset.Fields("Alt_SubscriptionType")
                        Alt_SubscriptionName.Value = Recordset.Fields("Alt_SubscriptionName")
                        Alt_SubscriptionDateTimeCreated.Value = Recordset.Fields("Alt_SubscriptionDateTimeCreated")
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    AltRowControls.Show
                Else
                    If HasNext Then
                        SubscriptionID.Value = Recordset.Fields("SubscriptionID")
                        SubscriptionID.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        SubscriptionID.Parameters = CCAddParam(SubscriptionID.Parameters, "SubscriptionID", Recordset.Fields("SubscriptionID_param1"))
                        SubscriptionID.Page = "AdminSubscriptionEdit.asp"
                        SubscriptionType.Value = Recordset.Fields("SubscriptionType")
                        SubscriptionName.Value = Recordset.Fields("SubscriptionName")
                        SubscriptionDateTimeCreated.Value = Recordset.Fields("SubscriptionDateTimeCreated")
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
'End Subscription Show Method

'Subscription PageSize Property Let @33-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End Subscription PageSize Property Let

'Subscription PageSize Property Get @33-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End Subscription PageSize Property Get

'Subscription RowNumber Property Get @33-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End Subscription RowNumber Property Get

'Subscription HasNextRow Function @33-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End Subscription HasNextRow Function

End Class 'End Subscription Class @33-A61BA892

Class clsSubscriptionDataSource 'SubscriptionDataSource Class @33-1E6C1D58

'DataSource Variables @33-05FDFBB7
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
    Public SubscriptionID
    Public SubscriptionID_param1
    Public SubscriptionType
    Public SubscriptionName
    Public SubscriptionDateTimeCreated
    Public Alt_SubscriptionID
    Public Alt_SubscriptionID_param1
    Public Alt_SubscriptionType
    Public Alt_SubscriptionName
    Public Alt_SubscriptionDateTimeCreated
'End DataSource Variables

'DataSource Class_Initialize Event @33-D6C11650
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set SubscriptionID = CCCreateField("SubscriptionID", "SubscriptionID", ccsInteger, Empty, Recordset)
        Set SubscriptionID_param1 = CCCreateField("SubscriptionID_param1", "SubscriptionID", ccsText, Empty, Recordset)
        Set SubscriptionType = CCCreateField("SubscriptionType", "SubscriptionType", ccsText, Empty, Recordset)
        Set SubscriptionName = CCCreateField("SubscriptionName", "SubscriptionName", ccsText, Empty, Recordset)
        Set SubscriptionDateTimeCreated = CCCreateField("SubscriptionDateTimeCreated", "SubscriptionDateTimeCreated", ccsDate, Array("yyyy", "-", "mm", "-", "dd", " ", "HH", ":", "nn", ":", "ss"), Recordset)
        Set Alt_SubscriptionID = CCCreateField("Alt_SubscriptionID", "SubscriptionID", ccsInteger, Empty, Recordset)
        Set Alt_SubscriptionID_param1 = CCCreateField("Alt_SubscriptionID_param1", "SubscriptionID", ccsText, Empty, Recordset)
        Set Alt_SubscriptionType = CCCreateField("Alt_SubscriptionType", "SubscriptionType", ccsText, Empty, Recordset)
        Set Alt_SubscriptionName = CCCreateField("Alt_SubscriptionName", "SubscriptionName", ccsText, Empty, Recordset)
        Set Alt_SubscriptionDateTimeCreated = CCCreateField("Alt_SubscriptionDateTimeCreated", "SubscriptionDateTimeCreated", ccsDate, Array("yyyy", "-", "mm", "-", "dd", " ", "HH", ":", "nn", ":", "ss"), Recordset)
        Fields.AddFields Array(SubscriptionID,  SubscriptionID_param1,  SubscriptionType,  SubscriptionName,  SubscriptionDateTimeCreated,  Alt_SubscriptionID,  Alt_SubscriptionID_param1, _
             Alt_SubscriptionType,  Alt_SubscriptionName,  Alt_SubscriptionDateTimeCreated)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_SubscriptionID", "SubscriptionID", ""), _
            Array("Sorter_SubscriptionType", "SubscriptionType", ""), _
            Array("Sorter_SubscriptionName", "SubscriptionName", ""), _
            Array("Sorter_SubscriptionDateTimeCreated", "SubscriptionDateTimeCreated", ""))

        SQL = "SELECT TOP {SqlParam_endRecord}  *  " & vbLf & _
        "FROM Subscription {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM Subscription"
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

'BuildTableWhere Method @33-B94EDC42
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
            .AddParameter 2, "urls_keyword", ccsInteger, Empty, Empty, Empty, False
            .AddParameter 3, "urls_keyword", ccsText, Empty, Empty, Empty, False
            .AddParameter 4, "urls_keyword", ccsText, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[SubscriptionSiteID]", .getParamByID(1))
            .Criterion(2) = .Operation(opContains, False, "[SubscriptionID]", .getParamByID(2))
            .Criterion(3) = .Operation(opContains, False, "[SubscriptionType]", .getParamByID(3))
            .Criterion(4) = .Operation(opContains, False, "[SubscriptionName]", .getParamByID(4))
            .AssembledWhere = .opAND(False, .Criterion(1), .opOR(True, .opOR(False, .Criterion(2), .Criterion(3)), .Criterion(4)))
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

End Class 'End SubscriptionDataSource Class @33-A61BA892


%>
