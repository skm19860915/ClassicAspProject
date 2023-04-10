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

'Initialize Page @1-AB9D1F48
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
Dim TrackingLinkCode
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminTrackingLinkGroupCodes.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminTrackingLinkGroupCodes.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-1240E005
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set TrackingLinkCode = New clsGridTrackingLinkCode
Menu.Value = DHTMLMenu

TrackingLinkCode.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminTrackingLinkGroupCodes_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

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

'Show Page @1-6F298D88
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, TrackingLinkCode))
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

'UnloadPage Sub @1-1D83FB10
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set TrackingLinkCode = Nothing
End Sub
'End UnloadPage Sub

Class clsGridTrackingLinkCode 'TrackingLinkCode Class @4-FA4A78EC

'TrackingLinkCode Variables @4-A555B9BB

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
    Dim Sorter_TrackingLinkCode
    Dim Sorter_TrackingLinkCodeDescription
    Dim TrackingLinkCode
    Dim TrackingLinkCodeDescription
    Dim ImageLink
    Dim Alt_TrackingLinkCode
    Dim Alt_TrackingLinkCodeDescription
    Dim Alt_ImageLink
    Dim TrackingLinkCode_TotalRecords
    Dim Navigator
'End TrackingLinkCode Variables

'TrackingLinkCode Class_Initialize Event @4-91F0A8C8
    Private Sub Class_Initialize()
        ComponentName = "TrackingLinkCode"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsTrackingLinkCodeDataSource
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
        ActiveSorter = CCGetParam("TrackingLinkCodeOrder", Empty)
        SortingDirection = CCGetParam("TrackingLinkCodeDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Link1 = CCCreateControl(ccsLink, "Link1", Empty, ccsText, Empty, CCGetRequestParam("Link1", ccsGet))
        Set Sorter_TrackingLinkCode = CCCreateSorter("Sorter_TrackingLinkCode", Me, FileName)
        Set Sorter_TrackingLinkCodeDescription = CCCreateSorter("Sorter_TrackingLinkCodeDescription", Me, FileName)
        Set TrackingLinkCode = CCCreateControl(ccsLabel, "TrackingLinkCode", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkCode", ccsGet))
        Set TrackingLinkCodeDescription = CCCreateControl(ccsLabel, "TrackingLinkCodeDescription", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkCodeDescription", ccsGet))
        Set ImageLink = CCCreateControl(ccsImageLink, "ImageLink", Empty, ccsText, Empty, CCGetRequestParam("ImageLink", ccsGet))
        Set Alt_TrackingLinkCode = CCCreateControl(ccsLabel, "Alt_TrackingLinkCode", Empty, ccsText, Empty, CCGetRequestParam("Alt_TrackingLinkCode", ccsGet))
        Set Alt_TrackingLinkCodeDescription = CCCreateControl(ccsLabel, "Alt_TrackingLinkCodeDescription", Empty, ccsText, Empty, CCGetRequestParam("Alt_TrackingLinkCodeDescription", ccsGet))
        Set Alt_ImageLink = CCCreateControl(ccsImageLink, "Alt_ImageLink", Empty, ccsText, Empty, CCGetRequestParam("Alt_ImageLink", ccsGet))
        Set TrackingLinkCode_TotalRecords = CCCreateControl(ccsLabel, "TrackingLinkCode_TotalRecords", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkCode_TotalRecords", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
    IsDSEmpty = True
    End Sub
'End TrackingLinkCode Class_Initialize Event

'TrackingLinkCode Initialize Method @4-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End TrackingLinkCode Initialize Method

'TrackingLinkCode Class_Terminate Event @4-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End TrackingLinkCode Class_Terminate Event

'TrackingLinkCode Show Method @4-F7C2B0B7
    Sub Show(Tpl)
        Dim HasNext
        If NOT Visible Then Exit Sub

        Dim RowBlock, SeparatorBlock, AltRowBlock

        With DataSource
            .Parameters("urlTrackingLinkGroupID") = CCGetRequestParam("TrackingLinkGroupID", ccsGET)
        End With

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        IsDSEmpty = Recordset.EOF

        Set TemplateBlock = Tpl.Block("Grid " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        Set RowBlock = TemplateBlock.Block("Row")
        Set SeparatorBlock = TemplateBlock.Block("Separator")
        Set AltRowBlock = TemplateBlock.Block("AltRow")
        Set StaticControls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(Link1, Sorter_TrackingLinkCode, Sorter_TrackingLinkCodeDescription, TrackingLinkCode_TotalRecords, Navigator))
            
            Link1.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
            Link1.Parameters = CCAddParam(Link1.Parameters, "TrackingLinkGroupID", CCGetRequestParam("TrackingLinkGroupID", ccsGET))
            Link1.Page = "AdminTrackingLinkCodeEdit.asp"
            
            Navigator.SetDataSource Recordset
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(TrackingLinkCode, TrackingLinkCodeDescription, ImageLink))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_TrackingLinkCode, Alt_TrackingLinkCodeDescription, Alt_ImageLink))

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
                        Alt_TrackingLinkCode.Value = Recordset.Fields("Alt_TrackingLinkCode")
                        Alt_TrackingLinkCodeDescription.Value = Recordset.Fields("Alt_TrackingLinkCodeDescription")
                        
                        Alt_ImageLink.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        Alt_ImageLink.Parameters = CCAddParam(Alt_ImageLink.Parameters, "TrackingLinkCodeID", Recordset.Fields("Alt_ImageLink_param1"))
                        Alt_ImageLink.Page = "AdminTrackingLinkCodeEdit.asp"
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    AltRowControls.Show
                Else
                    If HasNext Then
                        TrackingLinkCode.Value = Recordset.Fields("TrackingLinkCode")
                        TrackingLinkCodeDescription.Value = Recordset.Fields("TrackingLinkCodeDescription")
                        
                        ImageLink.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        ImageLink.Parameters = CCAddParam(ImageLink.Parameters, "TrackingLinkCodeID", Recordset.Fields("ImageLink_param1"))
                        ImageLink.Page = "AdminTrackingLinkCodeEdit.asp"
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    RowControls.Show
                End If
                RenderAltRow = NOT RenderAltRow
                If HasNext Then Recordset.MoveNext
                ShownRecords = ShownRecords + 1

                ' Parse Separator
                If NOT Recordset.EOF AND ShownRecords < PageSize Then _
                    SeparatorBlock.ParseTo ccsParseAccumulate, RowBlock
            Loop
            StaticControls.Show
        End If

    End Sub
'End TrackingLinkCode Show Method

'TrackingLinkCode PageSize Property Let @4-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End TrackingLinkCode PageSize Property Let

'TrackingLinkCode PageSize Property Get @4-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End TrackingLinkCode PageSize Property Get

'TrackingLinkCode RowNumber Property Get @4-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End TrackingLinkCode RowNumber Property Get

'TrackingLinkCode HasNextRow Function @4-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End TrackingLinkCode HasNextRow Function

End Class 'End TrackingLinkCode Class @4-A61BA892

Class clsTrackingLinkCodeDataSource 'TrackingLinkCodeDataSource Class @4-2803142A

'DataSource Variables @4-886438E1
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
    Public TrackingLinkCode
    Public TrackingLinkCodeDescription
    Public ImageLink_param1
    Public Alt_TrackingLinkCode
    Public Alt_TrackingLinkCodeDescription
    Public Alt_ImageLink_param1
'End DataSource Variables

'DataSource Class_Initialize Event @4-CF0A5F15
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set TrackingLinkCode = CCCreateField("TrackingLinkCode", "TrackingLinkCode", ccsText, Empty, Recordset)
        Set TrackingLinkCodeDescription = CCCreateField("TrackingLinkCodeDescription", "TrackingLinkCodeDescription", ccsText, Empty, Recordset)
        Set ImageLink_param1 = CCCreateField("ImageLink_param1", "TrackingLinkCodeID", ccsText, Empty, Recordset)
        Set Alt_TrackingLinkCode = CCCreateField("Alt_TrackingLinkCode", "TrackingLinkCode", ccsText, Empty, Recordset)
        Set Alt_TrackingLinkCodeDescription = CCCreateField("Alt_TrackingLinkCodeDescription", "TrackingLinkCodeDescription", ccsText, Empty, Recordset)
        Set Alt_ImageLink_param1 = CCCreateField("Alt_ImageLink_param1", "TrackingLinkCodeID", ccsText, Empty, Recordset)
        Fields.AddFields Array(TrackingLinkCode, TrackingLinkCodeDescription, Alt_TrackingLinkCode, Alt_TrackingLinkCodeDescription, ImageLink_param1, Alt_ImageLink_param1)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_TrackingLinkCode", "TrackingLinkCode", ""), _
            Array("Sorter_TrackingLinkCodeDescription", "TrackingLinkCodeDescription", ""))

        SQL = "SELECT TOP {SqlParam_endRecord}  *  " & vbLf & _
        "FROM TrackingLinkCode {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM TrackingLinkCode"
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

'BuildTableWhere Method @4-F1F40D83
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urlTrackingLinkGroupID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[TrackingLinkCodeTrackingLinkGroupID]", .getParamByID(1))
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

End Class 'End TrackingLinkCodeDataSource Class @4-A61BA892


%>
