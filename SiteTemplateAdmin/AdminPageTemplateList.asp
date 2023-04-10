<%@ CodePage=1252 %>
<%
'Include Common Files @1-C42A9701
%>
<!-- #INCLUDE FILE="Common.asp"-->
<!-- #INCLUDE FILE="Cache.asp" -->
<!-- #INCLUDE FILE="Template.asp" -->
<!-- #INCLUDE FILE="Sorter.asp" -->
<!-- #INCLUDE FILE="Navigator.asp" -->
<%
'End Include Common Files

'Initialize Page @1-296ACBC9
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
Dim PageTemplateSearch
Dim PageTemplate
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminPageTemplateList.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminPageTemplateList.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-3A73A510
CCSecurityRedirect "50", Empty
'End Authenticate User

'Initialize Objects @1-D1AB177B
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set PageTemplateSearch = new clsRecordPageTemplateSearch
Set PageTemplate = New clsGridPageTemplate
Menu.Value = DHTMLMenu

PageTemplate.Initialize DBSystem

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-B3BB3E6C
PageTemplateSearch.Operation
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

'Show Page @1-8EDE18D9
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, PageTemplateSearch, PageTemplate))
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

'UnloadPage Sub @1-134E829D
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set PageTemplateSearch = Nothing
    Set PageTemplate = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordPageTemplateSearch 'PageTemplateSearch Class @3-1688CCD9

'PageTemplateSearch Variables @3-6311E01D

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
    Dim s_PageTemplateSearchString
    Dim PageTemplatePageSize
    Dim Button_DoSearch
'End PageTemplateSearch Variables

'PageTemplateSearch Class_Initialize Event @3-9C42068F
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
            FormSubmitted = (OperationMode(0) = "PageTemplateSearch")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "PageTemplateSearch"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set s_PageTemplateSearchString = CCCreateControl(ccsTextBox, "s_PageTemplateSearchString", Empty, ccsText, Empty, CCGetRequestParam("s_PageTemplateSearchString", Method))
        Set PageTemplatePageSize = CCCreateList(ccsListBox, "PageTemplatePageSize", Empty, ccsText, CCGetRequestParam("PageTemplatePageSize", Method), Empty)
        Set PageTemplatePageSize.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("", "5", "10", "25", "100"), _
            Array("Select Value", "5", "10", "25", "100")))
        Set Button_DoSearch = CCCreateButton("Button_DoSearch", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(s_PageTemplateSearchString, PageTemplatePageSize)
    End Sub
'End PageTemplateSearch Class_Initialize Event

'PageTemplateSearch Class_Terminate Event @3-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End PageTemplateSearch Class_Terminate Event

'PageTemplateSearch Validate Method @3-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End PageTemplateSearch Validate Method

'PageTemplateSearch Operation Method @3-A62618C9
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoSearch"
            If Button_DoSearch.Pressed Then
                PressedButton = "Button_DoSearch"
            End If
        End If
        Redirect = "AdminPageTemplateList.asp"
        If Validate() Then
            If PressedButton = "Button_DoSearch" Then
                If NOT Button_DoSearch.OnClick() Then
                    Redirect = ""
                Else
                    Redirect = "AdminPageTemplateList.asp?" & CCGetQueryString("Form", Array(PressedButton, "ccsForm", "Button_DoSearch.x", "Button_DoSearch.y", "Button_DoSearch"))
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End PageTemplateSearch Operation Method

'PageTemplateSearch Show Method @3-FBCE06A3
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "PageTemplateSearch" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(s_PageTemplateSearchString, PageTemplatePageSize, Button_DoSearch))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors s_PageTemplateSearchString.Errors
            Errors.AddErrors PageTemplatePageSize.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End PageTemplateSearch Show Method

End Class 'End PageTemplateSearch Class @3-A61BA892

Class clsGridPageTemplate 'PageTemplate Class @21-AA08A231

'PageTemplate Variables @21-CD9DDB15

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
    Dim Sorter_PageTemplateID
    Dim Sorter_PageTemplateName
    Dim Sorter_PageTemplateNickname
    Dim Sorter_PageTemplatePageType
	Dim Sorter_PageTemplateUserLastUpdateDateTime
    Dim PageTemplateID
    Dim PageTemplateName
    Dim PageTemplateNickname
    Dim PageTemplatePageType
    Dim PageTemplateUserLastUpdateBy
    Dim PageTemplateUserLastUpdateDateTime
    Dim Alt_PageTemplateID
    Dim Alt_PageTemplateName
    Dim Alt_PageTemplateNickname
    Dim Alt_PageTemplatePageType
    Dim Alt_PageTemplateUserLastUpdateBy
    Dim Alt_PageTemplateUserLastUpdateDateTime
    Dim Navigator
'End PageTemplate Variables

'PageTemplate Class_Initialize Event @21-86976423
    Private Sub Class_Initialize()
        ComponentName = "PageTemplate"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsPageTemplateDataSource
        Set Command = New clsCommand
        PageSize = CCGetParam(ComponentName & "PageSize", Empty)
        If IsNumeric(PageSize) And Len(PageSize) > 0 Then
            If PageSize <= 0 Then Errors.AddError(CCSLocales.GetText("CCS_GridPageSizeError", Empty))
            If PageSize > 100 Then PageSize = 100
        End If
        If NOT IsNumeric(PageSize) OR IsEmpty(PageSize) Then _
            PageSize = 50 _
        Else _
            PageSize = CInt(PageSize)
        PageNumber = CCGetParam(ComponentName & "Page", 1)
        If Not IsNumeric(PageNumber) And Len(PageNumber) > 0 Then
            Errors.AddError(CCSLocales.GetText("CCS_GridPageNumberError", Empty))
            PageNumber = 1
        ElseIf Len(PageNumber) > 0 Then
            If PageNumber > 0 Then
                PageNumber = CInt(PageNumber)
            Else
                Errors.AddError(CCSLocales.GetText("CCS_GridPageNumberError", Empty))
                PageNumber = 1
            End If
        Else
            PageNumber = 1
        End If
        ActiveSorter = CCGetParam("PageTemplateOrder", Empty)
        SortingDirection = CCGetParam("PageTemplateDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Link1 = CCCreateControl(ccsLink, "Link1", Empty, ccsText, Empty, CCGetRequestParam("Link1", ccsGet))
        Set Sorter_PageTemplateID = CCCreateSorter("Sorter_PageTemplateID", Me, FileName)
        Set Sorter_PageTemplateName = CCCreateSorter("Sorter_PageTemplateName", Me, FileName)
        Set Sorter_PageTemplateNickname = CCCreateSorter("Sorter_PageTemplateNickname", Me, FileName)
        Set Sorter_PageTemplatePageType = CCCreateSorter("Sorter_PageTemplatePageType", Me, FileName)
		Set Sorter_PageTemplateUserLastUpdateDateTime = CCCreateSorter("Sorter_PageTemplateUserLastUpdateDateTime", Me, FileName)
        Set PageTemplateID = CCCreateControl(ccsLink, "PageTemplateID", Empty, ccsInteger, Empty, CCGetRequestParam("PageTemplateID", ccsGet))
        Set PageTemplateName = CCCreateControl(ccsLabel, "PageTemplateName", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateName", ccsGet))
        Set PageTemplateNickname = CCCreateControl(ccsLabel, "PageTemplateNickname", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateNickname", ccsGet))
        Set PageTemplatePageType = CCCreateControl(ccsLabel, "PageTemplatePageType", Empty, ccsText, Empty, CCGetRequestParam("PageTemplatePageType", ccsGet))
        Set PageTemplateUserLastUpdateBy = CCCreateControl(ccsLabel, "PageTemplateUserLastUpdateBy", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateUserLastUpdateBy", ccsGet))
        Set PageTemplateUserLastUpdateDateTime = CCCreateControl(ccsLabel, "PageTemplateUserLastUpdateDateTime", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateUserLastUpdateDateTime", ccsGet))
        Set Alt_PageTemplateID = CCCreateControl(ccsLink, "Alt_PageTemplateID", Empty, ccsInteger, Empty, CCGetRequestParam("Alt_PageTemplateID", ccsGet))
        Set Alt_PageTemplateName = CCCreateControl(ccsLabel, "Alt_PageTemplateName", Empty, ccsText, Empty, CCGetRequestParam("Alt_PageTemplateName", ccsGet))
        Set Alt_PageTemplateNickname = CCCreateControl(ccsLabel, "Alt_PageTemplateNickname", Empty, ccsText, Empty, CCGetRequestParam("Alt_PageTemplateNickname", ccsGet))
        Set Alt_PageTemplatePageType = CCCreateControl(ccsLabel, "Alt_PageTemplatePageType", Empty, ccsText, Empty, CCGetRequestParam("Alt_PageTemplatePageType", ccsGet))
        Set Alt_PageTemplateUserLastUpdateBy = CCCreateControl(ccsLabel, "Alt_PageTemplateUserLastUpdateBy", Empty, ccsText, Empty, CCGetRequestParam("Alt_PageTemplateUserLastUpdateBy", ccsGet))
        Set Alt_PageTemplateUserLastUpdateDateTime = CCCreateControl(ccsLabel, "Alt_PageTemplateUserLastUpdateDateTime", Empty, ccsText, Empty, CCGetRequestParam("Alt_PageTemplateUserLastUpdateDateTime", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
    IsDSEmpty = True
    End Sub
'End PageTemplate Class_Initialize Event

'PageTemplate Initialize Method @21-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End PageTemplate Initialize Method

'PageTemplate Class_Terminate Event @21-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End PageTemplate Class_Terminate Event

'PageTemplate Show Method @21-FBB803C0
    Sub Show(Tpl)
        Dim HasNext
        If NOT Visible Then Exit Sub

        Dim RowBlock, AltRowBlock

        With DataSource
            .Parameters("sesSiteID") = Session("SiteID")
            .Parameters("urls_PageTemplateSearchString") = CCGetRequestParam("s_PageTemplateSearchString", ccsGET)
        End With

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        If DataSource.Errors.Count = 0 Then IsDSEmpty = Recordset.EOF

        Set TemplateBlock = Tpl.Block("Grid " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        Set RowBlock = TemplateBlock.Block("Row")
        Set AltRowBlock = TemplateBlock.Block("AltRow")
        Set StaticControls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(Link1, Sorter_PageTemplateID, Sorter_PageTemplateName, Sorter_PageTemplateNickname, Sorter_PageTemplatePageType, Sorter_PageTemplateUserLastUpdateDateTime, Navigator))
            
            Link1.Parameters = CCGetQueryString("QueryString", Array("ccsForm"))
            Link1.Page = "AdminPageTemplateEdit.asp"
            Navigator.SetDataSource Recordset
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(PageTemplateID, PageTemplateName, PageTemplateNickname, PageTemplatePageType, PageTemplateUserLastUpdateBy, PageTemplateUserLastUpdateDateTime))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_PageTemplateID, Alt_PageTemplateName, Alt_PageTemplateNickname, Alt_PageTemplatePageType, Alt_PageTemplateUserLastUpdateBy, Alt_PageTemplateUserLastUpdateDateTime))

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
            HasNext = HasNextRow()
            ForceIteration = False
            Do While ForceIteration Or HasNext
                If RenderAltRow Then
                    If HasNext Then
                        Alt_PageTemplateID.Value = Recordset.Fields("Alt_PageTemplateID")
                        Alt_PageTemplateID.Link = ""
                        Alt_PageTemplateID.Parameters = CCAddParam(Alt_PageTemplateID.Parameters, "PageTemplateID", Recordset.Fields("Alt_PageTemplateID_param1"))
                        Alt_PageTemplateID.Page = "AdminPageTemplateEdit.asp"
                        Alt_PageTemplateName.Value = Recordset.Fields("Alt_PageTemplateName")
                        Alt_PageTemplateNickname.Value = Recordset.Fields("Alt_PageTemplateNickname")
                        Alt_PageTemplatePageType.Value = Recordset.Fields("Alt_PageTemplatePageType")
                        Alt_PageTemplateUserLastUpdateBy.Value = Recordset.Fields("Alt_PageTemplateUserLastUpdateBy")
                        Alt_PageTemplateUserLastUpdateDateTime.Value = Recordset.Fields("Alt_PageTemplateUserLastUpdateDateTime")
                    End If
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    AltRowControls.Show
                Else
                    If HasNext Then
                        PageTemplateID.Value = Recordset.Fields("PageTemplateID")
                        PageTemplateID.Link = ""
                        PageTemplateID.Parameters = CCAddParam(PageTemplateID.Parameters, "PageTemplateID", Recordset.Fields("PageTemplateID_param1"))
                        PageTemplateID.Page = "AdminPageTemplateEdit.asp"
                        PageTemplateName.Value = Recordset.Fields("PageTemplateName")
                        PageTemplateNickname.Value = Recordset.Fields("PageTemplateNickname")
                        PageTemplatePageType.Value = Recordset.Fields("PageTemplatePageType")
                        PageTemplateUserLastUpdateBy.Value = Recordset.Fields("PageTemplateUserLastUpdateBy")
                        PageTemplateUserLastUpdateDateTime.Value = Recordset.Fields("PageTemplateUserLastUpdateDateTime")
                    End If
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    RowControls.Show
                End If
                RenderAltRow = NOT RenderAltRow
                If HasNext Then Recordset.MoveNext
                ShownRecords = ShownRecords + 1
                HasNext = HasNextRow()
            Loop
            StaticControls.Show
        End If

    End Sub
'End PageTemplate Show Method

'PageTemplate PageSize Property Let @21-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End PageTemplate PageSize Property Let

'PageTemplate PageSize Property Get @21-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End PageTemplate PageSize Property Get

'PageTemplate RowNumber Property Get @21-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End PageTemplate RowNumber Property Get

'PageTemplate HasNextRow Function @21-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End PageTemplate HasNextRow Function

End Class 'End PageTemplate Class @21-A61BA892

Class clsPageTemplateDataSource 'PageTemplateDataSource Class @21-FB673E35

'DataSource Variables @21-0723BE21
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
    Public PageTemplateID
    Public PageTemplateID_param1
    Public PageTemplateName
    Public PageTemplateNickname
    Public PageTemplatePageType
    Public PageTemplateUserLastUpdateBy
    Public PageTemplateUserLastUpdateDateTime
    Public Alt_PageTemplateID
    Public Alt_PageTemplateID_param1
    Public Alt_PageTemplateName
    Public Alt_PageTemplateNickname
    Public Alt_PageTemplatePageType
    Public Alt_PageTemplateUserLastUpdateBy
    Public Alt_PageTemplateUserLastUpdateDateTime
'End DataSource Variables

'DataSource Class_Initialize Event @21-9A982A3A
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set PageTemplateID = CCCreateField("PageTemplateID", "PageTemplateID", ccsInteger, Empty, Recordset)
        Set PageTemplateID_param1 = CCCreateField("PageTemplateID_param1", "PageTemplateID", ccsText, Empty, Recordset)
        Set PageTemplateName = CCCreateField("PageTemplateName", "PageTemplateName", ccsText, Empty, Recordset)
        Set PageTemplateNickname = CCCreateField("PageTemplateNickname", "PageTemplateNickname", ccsText, Empty, Recordset)
        Set PageTemplatePageType = CCCreateField("PageTemplatePageType", "PageTemplatePageType", ccsText, Empty, Recordset)
        Set PageTemplateUserLastUpdateBy = CCCreateField("PageTemplateUserLastUpdateBy", "PageTemplateUserLastUpdateBy", ccsText, Empty, Recordset)
        Set PageTemplateUserLastUpdateDateTime = CCCreateField("PageTemplateUserLastUpdateDateTime", "PageTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Set Alt_PageTemplateID = CCCreateField("Alt_PageTemplateID", "PageTemplateID", ccsInteger, Empty, Recordset)
        Set Alt_PageTemplateID_param1 = CCCreateField("Alt_PageTemplateID_param1", "PageTemplateID", ccsText, Empty, Recordset)
        Set Alt_PageTemplateName = CCCreateField("Alt_PageTemplateName", "PageTemplateName", ccsText, Empty, Recordset)
        Set Alt_PageTemplateNickname = CCCreateField("Alt_PageTemplateNickname", "PageTemplateNickname", ccsText, Empty, Recordset)
        Set Alt_PageTemplatePageType = CCCreateField("Alt_PageTemplatePageType", "PageTemplatePageType", ccsText, Empty, Recordset)
        Set Alt_PageTemplateUserLastUpdateBy = CCCreateField("Alt_PageTemplateUserLastUpdateBy", "PageTemplateUserLastUpdateBy", ccsText, Empty, Recordset)
        Set Alt_PageTemplateUserLastUpdateDateTime = CCCreateField("Alt_PageTemplateUserLastUpdateDateTime", "PageTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Fields.AddFields Array(PageTemplateID,  PageTemplateID_param1,  PageTemplateName,  PageTemplateNickname,  PageTemplatePageType,  PageTemplateUserLastUpdateBy,  PageTemplateUserLastUpdateDateTime, _
             Alt_PageTemplateID,  Alt_PageTemplateID_param1,  Alt_PageTemplateName,  Alt_PageTemplateNickname,  Alt_PageTemplatePageType,  Alt_PageTemplateUserLastUpdateBy,  Alt_PageTemplateUserLastUpdateDateTime)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_PageTemplateID", "PageTemplateID", ""), _
            Array("Sorter_PageTemplateName", "PageTemplateName", ""), _
            Array("Sorter_PageTemplateNickname", "PageTemplateNickname", ""), _
            Array("Sorter_PageTemplatePageType", "PageTemplatePageType", ""), _
			Array("Sorter_PageTemplateUserLastUpdateDateTime", "PageTemplateUserLastUpdateDateTime", ""))
		If Request.QueryString("PageTemplateOrder") = "" then
        	SQL = "EXEC sp_getPageTemplateList {SiteID}, '{s_PageTemplateSearchString}', 'PageTemplate', 'PageTemplatePageType'"
		else
			SQL = "EXEC sp_getPageTemplateList {SiteID}, '{s_PageTemplateSearchString}', 'PageTemplate', '" & Mid(Request.QueryString("PageTemplateOrder"),8) & " " & Request.QueryString("PageTemplateDir") & "'"
		end if
        CountSQL = " " & vbLf & _
        "SELECT COUNT(*) FROM (EXEC sp_getPageTemplateList {SiteID}, '{s_PageTemplateSearchString}', 'PageTemplate') cnt"
        Where = ""
        Order = ""
        StaticOrder = ""
		'response.Write "<br><br>" & SQL & "<br>" & Request.QueryString("PageTemplateOrder")
    End Sub
'End DataSource Class_Initialize Event

'SetOrder Method @21-68FC9576
    Sub SetOrder(Column, Direction)
        Order = Recordset.GetOrder(Order, Column, Direction, Orders)
    End Sub
'End SetOrder Method

'BuildTableWhere Method @21-F0A2957C
    Public Sub BuildTableWhere()
        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter "SiteID", "sesSiteID", ccsInteger, Empty, Empty, -1, False
            .AddParameter "s_PageTemplateSearchString", "urls_PageTemplateSearchString", ccsText, Empty, Empty, Empty, False
        End With
    End Sub
'End BuildTableWhere Method

'Open Method @21-CA87DA7C
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
        Cmd.CommandType = dsSQL
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildSelect", Me)
        Cmd.SQL = SQL
        Cmd.CountSQL = CountSQL
        BuildTableWhere
        Set Cmd.WhereParameters = WhereParameters
        Cmd.Where = Where
        'Cmd.OrderBy = Order
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

'DataSource Class_Terminate Event @21-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

End Class 'End PageTemplateDataSource Class @21-A61BA892


%>
