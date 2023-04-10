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

'Initialize Page @1-3B2BA85B
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
Dim PageTemplateSectionSearch
Dim PageTemplateSection
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminPageTemplateSectionList.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminPageTemplateSectionList.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Initialize Objects @1-D23D2502
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set PageTemplateSectionSearch = new clsRecordPageTemplateSectionSearch
Set PageTemplateSection = New clsGridPageTemplateSection
Menu.Value = DHTMLMenu

PageTemplateSection.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminPageTemplateSectionList_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-067FF450
PageTemplateSectionSearch.Operation
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

'Show Page @1-6CC84DBB
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, PageTemplateSectionSearch, PageTemplateSection))
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

'UnloadPage Sub @1-E375A250
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set PageTemplateSectionSearch = Nothing
    Set PageTemplateSection = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordPageTemplateSectionSearch 'PageTemplateSectionSearch Class @34-2277B6BC

'PageTemplateSectionSearch Variables @34-693FD834

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
    Dim PageTemplateSectionPageSize
    Dim Button_DoSearch
'End PageTemplateSectionSearch Variables

'PageTemplateSectionSearch Class_Initialize Event @34-6C89F34F
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
            FormSubmitted = (OperationMode(0) = "PageTemplateSectionSearch")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "PageTemplateSectionSearch"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set s_keyword = CCCreateControl(ccsTextBox, "s_keyword", Empty, ccsText, Empty, CCGetRequestParam("s_keyword", Method))
        Set PageTemplateSectionPageSize = CCCreateList(ccsListBox, "PageTemplateSectionPageSize", Empty, ccsText, CCGetRequestParam("PageTemplateSectionPageSize", Method), Empty)
        Set PageTemplateSectionPageSize.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("", "5", "10", "25", "100"), _
            Array("Select Value", "5", "10", "25", "100")))
        Set Button_DoSearch = CCCreateButton("Button_DoSearch", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(s_keyword, PageTemplateSectionPageSize)
    End Sub
'End PageTemplateSectionSearch Class_Initialize Event

'PageTemplateSectionSearch Class_Terminate Event @34-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End PageTemplateSectionSearch Class_Terminate Event

'PageTemplateSectionSearch Validate Method @34-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End PageTemplateSectionSearch Validate Method

'PageTemplateSectionSearch Operation Method @34-4598FAE1
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoSearch"
            If Button_DoSearch.Pressed Then
                PressedButton = "Button_DoSearch"
            End If
        End If
        Redirect = "AdminPageTemplateSectionList.asp"
        If Validate() Then
            If PressedButton = "Button_DoSearch" Then
                If NOT Button_DoSearch.OnClick() Then
                    Redirect = ""
                Else
                    Redirect = "AdminPageTemplateSectionList.asp?" & CCGetQueryString("Form", Array(PressedButton, "ccsForm", "Button_DoSearch.x", "Button_DoSearch.y", "Button_DoSearch"))
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End PageTemplateSectionSearch Operation Method

'PageTemplateSectionSearch Show Method @34-7ACF9AD0
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "PageTemplateSectionSearch" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(s_keyword, PageTemplateSectionPageSize, Button_DoSearch))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors s_keyword.Errors
            Errors.AddErrors PageTemplateSectionPageSize.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End PageTemplateSectionSearch Show Method

End Class 'End PageTemplateSectionSearch Class @34-A61BA892

Class clsGridPageTemplateSection 'PageTemplateSection Class @33-AAA8EFD0

'PageTemplateSection Variables @33-FB0C2E26

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
    Dim PageTemplateSection_TotalRecords
    Dim Sorter_PageTemplateSectionID
    Dim Sorter_PageTemplateSectionName
    Dim PageTemplateSectionID
    Dim PageTemplateSectionName
    Dim Alt_PageTemplateSectionID
    Dim Alt_PageTemplateSectionName
    Dim Navigator
'End PageTemplateSection Variables

'PageTemplateSection Class_Initialize Event @33-6F85F91D
    Private Sub Class_Initialize()
        ComponentName = "PageTemplateSection"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsPageTemplateSectionDataSource
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
        ActiveSorter = CCGetParam("PageTemplateSectionOrder", Empty)
        SortingDirection = CCGetParam("PageTemplateSectionDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Link1 = CCCreateControl(ccsLink, "Link1", Empty, ccsText, Empty, CCGetRequestParam("Link1", ccsGet))
        Set PageTemplateSection_TotalRecords = CCCreateControl(ccsLabel, "PageTemplateSection_TotalRecords", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateSection_TotalRecords", ccsGet))
        Set Sorter_PageTemplateSectionID = CCCreateSorter("Sorter_PageTemplateSectionID", Me, FileName)
        Set Sorter_PageTemplateSectionName = CCCreateSorter("Sorter_PageTemplateSectionName", Me, FileName)
        Set PageTemplateSectionID = CCCreateControl(ccsLink, "PageTemplateSectionID", Empty, ccsInteger, Empty, CCGetRequestParam("PageTemplateSectionID", ccsGet))
        Set PageTemplateSectionName = CCCreateControl(ccsLabel, "PageTemplateSectionName", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateSectionName", ccsGet))
        Set Alt_PageTemplateSectionID = CCCreateControl(ccsLink, "Alt_PageTemplateSectionID", Empty, ccsInteger, Empty, CCGetRequestParam("Alt_PageTemplateSectionID", ccsGet))
        Set Alt_PageTemplateSectionName = CCCreateControl(ccsLabel, "Alt_PageTemplateSectionName", Empty, ccsText, Empty, CCGetRequestParam("Alt_PageTemplateSectionName", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
    IsDSEmpty = True
    End Sub
'End PageTemplateSection Class_Initialize Event

'PageTemplateSection Initialize Method @33-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End PageTemplateSection Initialize Method

'PageTemplateSection Class_Terminate Event @33-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End PageTemplateSection Class_Terminate Event

'PageTemplateSection Show Method @33-923F3270
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
            Array(Link1, PageTemplateSection_TotalRecords, Sorter_PageTemplateSectionID, Sorter_PageTemplateSectionName, Navigator))
            
            Link1.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
            Link1.Page = "AdminPageTemplateSectionEdit.asp"
            
            Navigator.SetDataSource Recordset
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(PageTemplateSectionID, PageTemplateSectionName))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_PageTemplateSectionID, Alt_PageTemplateSectionName))

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
                        Alt_PageTemplateSectionID.Value = Recordset.Fields("Alt_PageTemplateSectionID")
                        Alt_PageTemplateSectionID.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        Alt_PageTemplateSectionID.Parameters = CCAddParam(Alt_PageTemplateSectionID.Parameters, "PageTemplateSectionID", Recordset.Fields("Alt_PageTemplateSectionID_param1"))
                        Alt_PageTemplateSectionID.Page = "AdminPageTemplateSectionEdit.asp"
                        Alt_PageTemplateSectionName.Value = Recordset.Fields("Alt_PageTemplateSectionName")
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    AltRowControls.Show
                Else
                    If HasNext Then
                        PageTemplateSectionID.Value = Recordset.Fields("PageTemplateSectionID")
                        PageTemplateSectionID.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
                        PageTemplateSectionID.Parameters = CCAddParam(PageTemplateSectionID.Parameters, "PageTemplateSectionID", Recordset.Fields("PageTemplateSectionID_param1"))
                        PageTemplateSectionID.Page = "AdminPageTemplateSectionEdit.asp"
                        PageTemplateSectionName.Value = Recordset.Fields("PageTemplateSectionName")
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
'End PageTemplateSection Show Method

'PageTemplateSection PageSize Property Let @33-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End PageTemplateSection PageSize Property Let

'PageTemplateSection PageSize Property Get @33-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End PageTemplateSection PageSize Property Get

'PageTemplateSection RowNumber Property Get @33-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End PageTemplateSection RowNumber Property Get

'PageTemplateSection HasNextRow Function @33-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End PageTemplateSection HasNextRow Function

End Class 'End PageTemplateSection Class @33-A61BA892

Class clsPageTemplateSectionDataSource 'PageTemplateSectionDataSource Class @33-726CDD35

'DataSource Variables @33-DAE3F2C2
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
    Public PageTemplateSectionID
    Public PageTemplateSectionID_param1
    Public PageTemplateSectionName
    Public Alt_PageTemplateSectionID
    Public Alt_PageTemplateSectionID_param1
    Public Alt_PageTemplateSectionName
'End DataSource Variables

'DataSource Class_Initialize Event @33-665B6495
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set PageTemplateSectionID = CCCreateField("PageTemplateSectionID", "PageTemplateSectionID", ccsInteger, Empty, Recordset)
        Set PageTemplateSectionID_param1 = CCCreateField("PageTemplateSectionID_param1", "PageTemplateSectionID", ccsText, Empty, Recordset)
        Set PageTemplateSectionName = CCCreateField("PageTemplateSectionName", "PageTemplateSectionName", ccsText, Empty, Recordset)
        Set Alt_PageTemplateSectionID = CCCreateField("Alt_PageTemplateSectionID", "PageTemplateSectionID", ccsInteger, Empty, Recordset)
        Set Alt_PageTemplateSectionID_param1 = CCCreateField("Alt_PageTemplateSectionID_param1", "PageTemplateSectionID", ccsText, Empty, Recordset)
        Set Alt_PageTemplateSectionName = CCCreateField("Alt_PageTemplateSectionName", "PageTemplateSectionName", ccsText, Empty, Recordset)
        Fields.AddFields Array(PageTemplateSectionID, PageTemplateSectionID_param1, PageTemplateSectionName, Alt_PageTemplateSectionID, Alt_PageTemplateSectionID_param1, Alt_PageTemplateSectionName)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_PageTemplateSectionID", "PageTemplateSectionID", ""), _
            Array("Sorter_PageTemplateSectionName", "PageTemplateSectionName", ""))

        SQL = "SELECT TOP {SqlParam_endRecord}  *  " & vbLf & _
        "FROM PageTemplateSection {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM PageTemplateSection"
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

'BuildTableWhere Method @33-2336CC9C
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
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplateSectionSiteID]", .getParamByID(1))
            .Criterion(2) = .Operation(opContains, False, "[PageTemplateSectionID]", .getParamByID(2))
            .Criterion(3) = .Operation(opContains, False, "[PageTemplateSectionName]", .getParamByID(3))
            .AssembledWhere = .opAND(False, .Criterion(1), .opOR(True, .Criterion(2), .Criterion(3)))
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

End Class 'End PageTemplateSectionDataSource Class @33-A61BA892


%>
