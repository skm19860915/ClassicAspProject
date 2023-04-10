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

'Initialize Page @1-B7E91D19
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
Dim TrackingLinkGroupSearch
Dim TrackingLinkGroup
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminTrackingLinkGroupList.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminTrackingLinkGroupList.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-A58833D8
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set TrackingLinkGroupSearch = new clsRecordTrackingLinkGroupSearch
Set TrackingLinkGroup = New clsGridTrackingLinkGroup
Menu.Value = DHTMLMenu

TrackingLinkGroup.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminTrackingLinkGroupList_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-8365879D
TrackingLinkGroupSearch.Operation
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

'Show Page @1-555A8980
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, TrackingLinkGroupSearch, TrackingLinkGroup))
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

'UnloadPage Sub @1-9248EA07
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set TrackingLinkGroupSearch = Nothing
    Set TrackingLinkGroup = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordTrackingLinkGroupSearch 'TrackingLinkGroupSearch Class @5-386B1D1F

'TrackingLinkGroupSearch Variables @5-0E9A7FAF

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
    Dim s_TrackingLinkGroupName
    Dim TrackingLinkGroupPageSize
    Dim Button_DoSearch
'End TrackingLinkGroupSearch Variables

'TrackingLinkGroupSearch Class_Initialize Event @5-A1F3A560
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
            FormSubmitted = (OperationMode(0) = "TrackingLinkGroupSearch")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "TrackingLinkGroupSearch"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set s_TrackingLinkGroupName = CCCreateControl(ccsTextBox, "s_TrackingLinkGroupName", Empty, ccsText, Empty, CCGetRequestParam("s_TrackingLinkGroupName", Method))
        Set TrackingLinkGroupPageSize = CCCreateList(ccsListBox, "TrackingLinkGroupPageSize", Empty, ccsText, CCGetRequestParam("TrackingLinkGroupPageSize", Method), Empty)
        Set TrackingLinkGroupPageSize.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("", "5", "10", "25", "100"), _
            Array("Select Value", "5", "10", "25", "100")))
        Set Button_DoSearch = CCCreateButton("Button_DoSearch", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(s_TrackingLinkGroupName, TrackingLinkGroupPageSize)
    End Sub
'End TrackingLinkGroupSearch Class_Initialize Event

'TrackingLinkGroupSearch Class_Terminate Event @5-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End TrackingLinkGroupSearch Class_Terminate Event

'TrackingLinkGroupSearch Validate Method @5-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End TrackingLinkGroupSearch Validate Method

'TrackingLinkGroupSearch Operation Method @5-E22A35E9
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoSearch"
            If Button_DoSearch.Pressed Then
                PressedButton = "Button_DoSearch"
            End If
        End If
        Redirect = "AdminTrackingList.asp"
        If Validate() Then
            If PressedButton = "Button_DoSearch" Then
                If NOT Button_DoSearch.OnClick() Then
                    Redirect = ""
                Else
                    Redirect = "AdminTrackingList.asp?" & CCGetQueryString("Form", Array(PressedButton, "ccsForm", "Button_DoSearch.x", "Button_DoSearch.y", "Button_DoSearch"))
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End TrackingLinkGroupSearch Operation Method

'TrackingLinkGroupSearch Show Method @5-134EB787
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "TrackingLinkGroupSearch" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(s_TrackingLinkGroupName, TrackingLinkGroupPageSize, Button_DoSearch))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors s_TrackingLinkGroupName.Errors
            Errors.AddErrors TrackingLinkGroupPageSize.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End TrackingLinkGroupSearch Show Method

End Class 'End TrackingLinkGroupSearch Class @5-A61BA892

Class clsGridTrackingLinkGroup 'TrackingLinkGroup Class @4-3A25AEA8

'TrackingLinkGroup Variables @4-F7C23B2A

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
    Dim Link2
    Dim Sorter_TrackingLinkGroupID
    Dim Sorter_TrackingLinkGroupName
    Dim TrackingLinkGroupID
    Dim TrackingLinkGroupName
    Dim ImageLink1
    Dim Alt_TrackingLinkGroupID
    Dim Alt_TrackingLinkGroupName
    Dim ImageLink2
    Dim TrackingLinkGroup_TotalRecords
    Dim Navigator
'End TrackingLinkGroup Variables

'TrackingLinkGroup Class_Initialize Event @4-9BE77D9D
    Private Sub Class_Initialize()
        ComponentName = "TrackingLinkGroup"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsTrackingLinkGroupDataSource
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
        ActiveSorter = CCGetParam("TrackingLinkGroupOrder", Empty)
        SortingDirection = CCGetParam("TrackingLinkGroupDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Link1 = CCCreateControl(ccsLink, "Link1", Empty, ccsText, Empty, CCGetRequestParam("Link1", ccsGet))
        Set Link2 = CCCreateControl(ccsLink, "Link2", Empty, ccsText, Empty, CCGetRequestParam("Link2", ccsGet))
        Set Sorter_TrackingLinkGroupID = CCCreateSorter("Sorter_TrackingLinkGroupID", Me, FileName)
        Set Sorter_TrackingLinkGroupName = CCCreateSorter("Sorter_TrackingLinkGroupName", Me, FileName)
        Set TrackingLinkGroupID = CCCreateControl(ccsLabel, "TrackingLinkGroupID", Empty, ccsInteger, Empty, CCGetRequestParam("TrackingLinkGroupID", ccsGet))
        Set TrackingLinkGroupName = CCCreateControl(ccsLink, "TrackingLinkGroupName", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkGroupName", ccsGet))
        Set ImageLink1 = CCCreateControl(ccsImageLink, "ImageLink1", Empty, ccsText, Empty, CCGetRequestParam("ImageLink1", ccsGet))
        Set Alt_TrackingLinkGroupID = CCCreateControl(ccsLabel, "Alt_TrackingLinkGroupID", Empty, ccsInteger, Empty, CCGetRequestParam("Alt_TrackingLinkGroupID", ccsGet))
        Set Alt_TrackingLinkGroupName = CCCreateControl(ccsLink, "Alt_TrackingLinkGroupName", Empty, ccsText, Empty, CCGetRequestParam("Alt_TrackingLinkGroupName", ccsGet))
        Set ImageLink2 = CCCreateControl(ccsImageLink, "ImageLink2", Empty, ccsText, Empty, CCGetRequestParam("ImageLink2", ccsGet))
        Set TrackingLinkGroup_TotalRecords = CCCreateControl(ccsLabel, "TrackingLinkGroup_TotalRecords", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkGroup_TotalRecords", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
    IsDSEmpty = True
    End Sub
'End TrackingLinkGroup Class_Initialize Event

'TrackingLinkGroup Initialize Method @4-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End TrackingLinkGroup Initialize Method

'TrackingLinkGroup Class_Terminate Event @4-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End TrackingLinkGroup Class_Terminate Event

'TrackingLinkGroup Show Method @4-B19244E7
    Sub Show(Tpl)
        Dim HasNext
        If NOT Visible Then Exit Sub

        Dim RowBlock, AltRowBlock

        With DataSource
            .Parameters("urls_TrackingLinkGroupName") = CCGetRequestParam("s_TrackingLinkGroupName", ccsGET)
            .Parameters("expr23") = 0
            .Parameters("sesSiteID") = Session("SiteID")
        End With

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        IsDSEmpty = Recordset.EOF

        Set TemplateBlock = Tpl.Block("Grid " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        Set RowBlock = TemplateBlock.Block("Row")
        Set AltRowBlock = TemplateBlock.Block("AltRow")
        Set StaticControls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(Link1, Link2, Sorter_TrackingLinkGroupID, Sorter_TrackingLinkGroupName, TrackingLinkGroup_TotalRecords, Navigator))
            
            Link1.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
            Link1.Page = "AdminTrackingLinkGroupEdit.asp"
            
            Link2.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
            Link2.Page = "AdminTrackingLinkCodeEdit.asp"
            
            Navigator.SetDataSource Recordset
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(TrackingLinkGroupID, TrackingLinkGroupName, ImageLink1))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_TrackingLinkGroupID, Alt_TrackingLinkGroupName, ImageLink2))

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
                        Alt_TrackingLinkGroupID.Value = Recordset.Fields("Alt_TrackingLinkGroupID")
                        Alt_TrackingLinkGroupName.Value = Recordset.Fields("Alt_TrackingLinkGroupName")
                        Alt_TrackingLinkGroupName.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        Alt_TrackingLinkGroupName.Parameters = CCAddParam(Alt_TrackingLinkGroupName.Parameters, "TrackingLinkGroupID", Recordset.Fields("Alt_TrackingLinkGroupName_param1"))
                        Alt_TrackingLinkGroupName.Page = "AdminTrackingLinkGroupCodes.asp"
                        
                        ImageLink2.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        ImageLink2.Parameters = CCAddParam(ImageLink2.Parameters, "TrackingLinkGroupID", Recordset.Fields("ImageLink2_param1"))
                        ImageLink2.Page = "AdminTrackingLinkGroupEdit.asp"
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    AltRowControls.Show
                Else
                    If HasNext Then
                        TrackingLinkGroupID.Value = Recordset.Fields("TrackingLinkGroupID")
                        TrackingLinkGroupName.Value = Recordset.Fields("TrackingLinkGroupName")
                        TrackingLinkGroupName.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        TrackingLinkGroupName.Parameters = CCAddParam(TrackingLinkGroupName.Parameters, "TrackingLinkGroupID", Recordset.Fields("TrackingLinkGroupName_param1"))
                        TrackingLinkGroupName.Page = "AdminTrackingLinkGroupCodes.asp"
                        
                        ImageLink1.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        ImageLink1.Parameters = CCAddParam(ImageLink1.Parameters, "TrackingLinkGroupID", Recordset.Fields("ImageLink1_param1"))
                        ImageLink1.Page = "AdminTrackingLinkGroupEdit.asp"
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
'End TrackingLinkGroup Show Method

'TrackingLinkGroup PageSize Property Let @4-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End TrackingLinkGroup PageSize Property Let

'TrackingLinkGroup PageSize Property Get @4-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End TrackingLinkGroup PageSize Property Get

'TrackingLinkGroup RowNumber Property Get @4-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End TrackingLinkGroup RowNumber Property Get

'TrackingLinkGroup HasNextRow Function @4-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End TrackingLinkGroup HasNextRow Function

End Class 'End TrackingLinkGroup Class @4-A61BA892

Class clsTrackingLinkGroupDataSource 'TrackingLinkGroupDataSource Class @4-D71AA91E

'DataSource Variables @4-E1FAAD68
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
    Public TrackingLinkGroupID
    Public TrackingLinkGroupID_param1
    Public TrackingLinkGroupName
    Public TrackingLinkGroupName_param1
    Public ImageLink1_param1
    Public Alt_TrackingLinkGroupID
    Public Alt_TrackingLinkGroupID_param1
    Public Alt_TrackingLinkGroupName
    Public Alt_TrackingLinkGroupName_param1
    Public ImageLink2_param1
'End DataSource Variables

'DataSource Class_Initialize Event @4-4F4ECBAC
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set TrackingLinkGroupID = CCCreateField("TrackingLinkGroupID", "TrackingLinkGroupID", ccsInteger, Empty, Recordset)
        Set TrackingLinkGroupID_param1 = CCCreateField("TrackingLinkGroupID_param1", "TrackingLinkGroupID", ccsText, Empty, Recordset)
        Set TrackingLinkGroupName = CCCreateField("TrackingLinkGroupName", "TrackingLinkGroupName", ccsText, Empty, Recordset)
        Set TrackingLinkGroupName_param1 = CCCreateField("TrackingLinkGroupName_param1", "TrackingLinkGroupID", ccsText, Empty, Recordset)
        Set ImageLink1_param1 = CCCreateField("ImageLink1_param1", "TrackingLinkGroupID", ccsText, Empty, Recordset)
        Set Alt_TrackingLinkGroupID = CCCreateField("Alt_TrackingLinkGroupID", "TrackingLinkGroupID", ccsInteger, Empty, Recordset)
        Set Alt_TrackingLinkGroupID_param1 = CCCreateField("Alt_TrackingLinkGroupID_param1", "TrackingLinkGroupID", ccsText, Empty, Recordset)
        Set Alt_TrackingLinkGroupName = CCCreateField("Alt_TrackingLinkGroupName", "TrackingLinkGroupName", ccsText, Empty, Recordset)
        Set Alt_TrackingLinkGroupName_param1 = CCCreateField("Alt_TrackingLinkGroupName_param1", "TrackingLinkGroupID", ccsText, Empty, Recordset)
        Set ImageLink2_param1 = CCCreateField("ImageLink2_param1", "TrackingLinkGroupID", ccsText, Empty, Recordset)
        Fields.AddFields Array(TrackingLinkGroupID,  TrackingLinkGroupID_param1,  TrackingLinkGroupName,  TrackingLinkGroupName_param1,  Alt_TrackingLinkGroupID,  Alt_TrackingLinkGroupID_param1,  Alt_TrackingLinkGroupName, _
             Alt_TrackingLinkGroupName_param1,  ImageLink1_param1,  ImageLink2_param1)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_TrackingLinkGroupID", "TrackingLinkGroupID", ""), _
            Array("Sorter_TrackingLinkGroupName", "TrackingLinkGroupName", ""))

        SQL = "SELECT TOP {SqlParam_endRecord}  *  " & vbLf & _
        "FROM TrackingLinkGroup {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM TrackingLinkGroup"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'SetOrder Method @4-68FC9576
    Sub SetOrder(Column, Direction)
        Order = Recordset.GetOrder(Order, Column, Direction, Orders)
    End Sub
'End SetOrder Method

'BuildTableWhere Method @4-050A7F17
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urls_TrackingLinkGroupName", ccsText, Empty, Empty, Empty, False
            .AddParameter 2, "expr23", ccsInteger, Empty, Empty, Empty, False
            .AddParameter 3, "sesSiteID", ccsInteger, Empty, Empty, -1, False
            .Criterion(1) = .Operation(opContains, False, "[TrackingLinkGroupName]", .getParamByID(1))
            .Criterion(2) = .Operation(opGreaterThan, False, "[TrackingLinkGroupID]", .getParamByID(2))
            .Criterion(3) = .Operation(opEqual, False, "[TrackingLinkGroupSiteID]", .getParamByID(3))
            .AssembledWhere = .opAND(False, .opAND(False, .Criterion(1), .Criterion(2)), .Criterion(3))
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

'Open Method @4-40984FC5
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

'DataSource Class_Terminate Event @4-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

End Class 'End TrackingLinkGroupDataSource Class @4-A61BA892


%>
