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

'Initialize Page @1-F2A044BF
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
Dim EmailTemplate1
Dim EmailTemplate
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminEmailTemplateArchiveList.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminEmailTemplateArchiveList.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-312296A5
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set EmailTemplate1 = new clsRecordEmailTemplate1
Set EmailTemplate = New clsGridEmailTemplate
Menu.Value = DHTMLMenu

EmailTemplate.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminEmailTemplateArchiveList_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-92CF0028
EmailTemplate1.Operation
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

'Show Page @1-AD8DEF1B
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, EmailTemplate1, EmailTemplate))
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

'UnloadPage Sub @1-6298A84A
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set EmailTemplate1 = Nothing
    Set EmailTemplate = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordEmailTemplate1 'EmailTemplate1 Class @39-CCCC6903

'EmailTemplate1 Variables @39-63F8D6A1

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
    Dim EmailTemplatePageSize
    Dim Button_DoSearch
'End EmailTemplate1 Variables

'EmailTemplate1 Class_Initialize Event @39-092661B9
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
            FormSubmitted = (OperationMode(0) = "EmailTemplate1")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "EmailTemplate1"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set s_keyword = CCCreateControl(ccsTextBox, "s_keyword", Empty, ccsText, Empty, CCGetRequestParam("s_keyword", Method))
        Set EmailTemplatePageSize = CCCreateList(ccsListBox, "EmailTemplatePageSize", Empty, ccsText, CCGetRequestParam("EmailTemplatePageSize", Method), Empty)
        Set EmailTemplatePageSize.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("", "5", "10", "25", "100"), _
            Array("Select Value", "5", "10", "25", "100")))
        Set Button_DoSearch = CCCreateButton("Button_DoSearch", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(s_keyword, EmailTemplatePageSize)
    End Sub
'End EmailTemplate1 Class_Initialize Event

'EmailTemplate1 Class_Terminate Event @39-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End EmailTemplate1 Class_Terminate Event

'EmailTemplate1 Validate Method @39-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End EmailTemplate1 Validate Method

'EmailTemplate1 Operation Method @39-D47A3335
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoSearch"
            If Button_DoSearch.Pressed Then
                PressedButton = "Button_DoSearch"
            End If
        End If
        Redirect = "AdminEmailTemplateArchiveList.asp"
        If Validate() Then
            If PressedButton = "Button_DoSearch" Then
                If NOT Button_DoSearch.OnClick() Then
                    Redirect = ""
                Else
                    Redirect = "AdminEmailTemplateArchiveList.asp?" & CCGetQueryString("Form", Array(PressedButton, "ccsForm", "Button_DoSearch.x", "Button_DoSearch.y", "Button_DoSearch"))
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End EmailTemplate1 Operation Method

'EmailTemplate1 Show Method @39-FF6C2719
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "EmailTemplate1" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(s_keyword, EmailTemplatePageSize, Button_DoSearch))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors s_keyword.Errors
            Errors.AddErrors EmailTemplatePageSize.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End EmailTemplate1 Show Method

End Class 'End EmailTemplate1 Class @39-A61BA892

Class clsGridEmailTemplate 'EmailTemplate Class @2-AE628E99

'EmailTemplate Variables @2-B874FC87

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
    Dim Sorter_EmailID
    Dim Sorter_EmailName
    Dim Sorter_EmailToAddress
    Dim Sorter_EmailFromAddress
    Dim Sorter_EmailSubject
    Dim EmailTemplateID
    Dim EmailTemplateName
    Dim EmailTemplateToAddress
    Dim EmailTemplateFromAddress
    Dim EmailTemplateSubject
    Dim EmailTemplateEmailType
    Dim EmailTemplateUserLastUpdateBy
    Dim EmailTemplateUserLastUpdateDateTime
    Dim Alt_EmailTemplateID
    Dim Alt_EmailTemplateName
    Dim Alt_EmailTemplateToAddress
    Dim Alt_EmailTemplateFromAddress
    Dim Alt_EmailTemplateSubject
    Dim Alt_EmailTemplateEmailType
    Dim Alt_EmailTemplateUserLastUpdateBy
    Dim Alt_EmailTemplateUserLastUpdateDateTime
    Dim Navigator
    Dim EmailTemplate_TotalRecords
'End EmailTemplate Variables

'EmailTemplate Class_Initialize Event @2-EA2F659C
    Private Sub Class_Initialize()
        ComponentName = "EmailTemplate"
        Visible = True
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        RenderAltRow = False
        Set Errors = New clsErrors
        Set DataSource = New clsEmailTemplateDataSource
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
        ActiveSorter = CCGetParam("EmailTemplateOrder", Empty)
        SortingDirection = CCGetParam("EmailTemplateDir", Empty)
        If NOT(SortingDirection = "ASC" OR SortingDirection = "DESC") Then _
            SortingDirection = Empty

        Set Sorter_EmailID = CCCreateSorter("Sorter_EmailID", Me, FileName)
        Set Sorter_EmailName = CCCreateSorter("Sorter_EmailName", Me, FileName)
        Set Sorter_EmailToAddress = CCCreateSorter("Sorter_EmailToAddress", Me, FileName)
        Set Sorter_EmailFromAddress = CCCreateSorter("Sorter_EmailFromAddress", Me, FileName)
        Set Sorter_EmailSubject = CCCreateSorter("Sorter_EmailSubject", Me, FileName)
        Set EmailTemplateID = CCCreateControl(ccsLink, "EmailTemplateID", Empty, ccsInteger, Empty, CCGetRequestParam("EmailTemplateID", ccsGet))
        Set EmailTemplateName = CCCreateControl(ccsLabel, "EmailTemplateName", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateName", ccsGet))
        Set EmailTemplateToAddress = CCCreateControl(ccsLabel, "EmailTemplateToAddress", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateToAddress", ccsGet))
        Set EmailTemplateFromAddress = CCCreateControl(ccsLabel, "EmailTemplateFromAddress", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateFromAddress", ccsGet))
        Set EmailTemplateSubject = CCCreateControl(ccsLabel, "EmailTemplateSubject", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateSubject", ccsGet))
        Set EmailTemplateEmailType = CCCreateControl(ccsLabel, "EmailTemplateEmailType", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateEmailType", ccsGet))
        Set EmailTemplateUserLastUpdateBy = CCCreateControl(ccsLabel, "EmailTemplateUserLastUpdateBy", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateUserLastUpdateBy", ccsGet))
        Set EmailTemplateUserLastUpdateDateTime = CCCreateControl(ccsLabel, "EmailTemplateUserLastUpdateDateTime", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateUserLastUpdateDateTime", ccsGet))
        Set Alt_EmailTemplateID = CCCreateControl(ccsLink, "Alt_EmailTemplateID", Empty, ccsInteger, Empty, CCGetRequestParam("Alt_EmailTemplateID", ccsGet))
        Set Alt_EmailTemplateName = CCCreateControl(ccsLabel, "Alt_EmailTemplateName", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateName", ccsGet))
        Set Alt_EmailTemplateToAddress = CCCreateControl(ccsLabel, "Alt_EmailTemplateToAddress", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateToAddress", ccsGet))
        Set Alt_EmailTemplateFromAddress = CCCreateControl(ccsLabel, "Alt_EmailTemplateFromAddress", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateFromAddress", ccsGet))
        Set Alt_EmailTemplateSubject = CCCreateControl(ccsLabel, "Alt_EmailTemplateSubject", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateSubject", ccsGet))
        Set Alt_EmailTemplateEmailType = CCCreateControl(ccsLabel, "Alt_EmailTemplateEmailType", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateEmailType", ccsGet))
        Set Alt_EmailTemplateUserLastUpdateBy = CCCreateControl(ccsLabel, "Alt_EmailTemplateUserLastUpdateBy", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateUserLastUpdateBy", ccsGet))
        Set Alt_EmailTemplateUserLastUpdateDateTime = CCCreateControl(ccsLabel, "Alt_EmailTemplateUserLastUpdateDateTime", Empty, ccsText, Empty, CCGetRequestParam("Alt_EmailTemplateUserLastUpdateDateTime", ccsGet))
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpCentered)
        Set EmailTemplate_TotalRecords = CCCreateControl(ccsLabel, "EmailTemplate_TotalRecords", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplate_TotalRecords", ccsGet))
    IsDSEmpty = True
    End Sub
'End EmailTemplate Class_Initialize Event

'EmailTemplate Initialize Method @2-2AEA3975
    Sub Initialize(objConnection)
        If NOT Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        DataSource.PageSize = PageSize
        DataSource.SetOrder ActiveSorter, SortingDirection
        DataSource.AbsolutePage = PageNumber
    End Sub
'End EmailTemplate Initialize Method

'EmailTemplate Class_Terminate Event @2-2C3914FE
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set DataSource = Nothing
        Set Command = Nothing
        Set Errors = Nothing
    End Sub
'End EmailTemplate Class_Terminate Event

'EmailTemplate Show Method @2-55A294FA
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
            Array(Sorter_EmailID, Sorter_EmailName, Sorter_EmailToAddress, Sorter_EmailFromAddress, Sorter_EmailSubject, Navigator, EmailTemplate_TotalRecords))
            Navigator.SetDataSource Recordset
            
        Set RowControls = CCCreateCollection(RowBlock, Null, ccsParseAccumulate, _
            Array(EmailTemplateID, EmailTemplateName, EmailTemplateToAddress, EmailTemplateFromAddress, EmailTemplateSubject, EmailTemplateEmailType, EmailTemplateUserLastUpdateBy, EmailTemplateUserLastUpdateDateTime))
        Set AltRowControls = CCCreateCollection(AltRowBlock, RowBlock, ccsParseAccumulate, _
            Array(Alt_EmailTemplateID, Alt_EmailTemplateName, Alt_EmailTemplateToAddress, Alt_EmailTemplateFromAddress, Alt_EmailTemplateSubject, Alt_EmailTemplateEmailType, Alt_EmailTemplateUserLastUpdateBy, Alt_EmailTemplateUserLastUpdateDateTime))

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
                        Alt_EmailTemplateID.Value = Recordset.Fields("Alt_EmailTemplateID")
                        Alt_EmailTemplateID.Link = ""
                        Alt_EmailTemplateID.Parameters = CCAddParam(Alt_EmailTemplateID.Parameters, "EmailTemplateID", Recordset.Fields("Alt_EmailTemplateID_param1"))
                        Alt_EmailTemplateID.Parameters = CCAddParam(Alt_EmailTemplateID.Parameters, "EmailTemplateUserLastUpdateDateTime", Recordset.Fields("Alt_EmailTemplateID_param2"))
                        Alt_EmailTemplateID.Page = "AdminEmailTemplateArchiveView.asp"
                        Alt_EmailTemplateName.Value = Recordset.Fields("Alt_EmailTemplateName")
                        Alt_EmailTemplateToAddress.Value = Recordset.Fields("Alt_EmailTemplateToAddress")
                        Alt_EmailTemplateFromAddress.Value = Recordset.Fields("Alt_EmailTemplateFromAddress")
                        Alt_EmailTemplateSubject.Value = Recordset.Fields("Alt_EmailTemplateSubject")
                        Alt_EmailTemplateEmailType.Value = Recordset.Fields("Alt_EmailTemplateEmailType")
                        Alt_EmailTemplateUserLastUpdateBy.Value = Recordset.Fields("Alt_EmailTemplateUserLastUpdateBy")
                        Alt_EmailTemplateUserLastUpdateDateTime.Value = Recordset.Fields("Alt_EmailTemplateUserLastUpdateDateTime")
                    End If
                    ForceIteration = HasNext
                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    If Not ForceIteration Then Exit Do
                    AltRowControls.Show
                Else
                    If HasNext Then
                        EmailTemplateID.Value = Recordset.Fields("EmailTemplateID")
                        EmailTemplateID.Link = ""
                        EmailTemplateID.Parameters = CCAddParam(EmailTemplateID.Parameters, "EmailTemplateID", Recordset.Fields("EmailTemplateID_param1"))
                        EmailTemplateID.Parameters = CCAddParam(EmailTemplateID.Parameters, "EmailTemplateUserLastUpdateDateTime", Recordset.Fields("EmailTemplateID_param2"))
                        EmailTemplateID.Page = "AdminEmailTemplateArchiveView.asp"
                        EmailTemplateName.Value = Recordset.Fields("EmailTemplateName")
                        EmailTemplateToAddress.Value = Recordset.Fields("EmailTemplateToAddress")
                        EmailTemplateFromAddress.Value = Recordset.Fields("EmailTemplateFromAddress")
                        EmailTemplateSubject.Value = Recordset.Fields("EmailTemplateSubject")
                        EmailTemplateEmailType.Value = Recordset.Fields("EmailTemplateEmailType")
                        EmailTemplateUserLastUpdateBy.Value = Recordset.Fields("EmailTemplateUserLastUpdateBy")
                        EmailTemplateUserLastUpdateDateTime.Value = Recordset.Fields("EmailTemplateUserLastUpdateDateTime")
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
'End EmailTemplate Show Method

'EmailTemplate PageSize Property Let @2-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End EmailTemplate PageSize Property Let

'EmailTemplate PageSize Property Get @2-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End EmailTemplate PageSize Property Get

'EmailTemplate RowNumber Property Get @2-F32EE2C6
    Public Property Get RowNumber()
        RowNumber = ShownRecords + 1
    End Property
'End EmailTemplate RowNumber Property Get

'EmailTemplate HasNextRow Function @2-9BECE27A
    Public Function HasNextRow()
        HasNextRow = NOT Recordset.EOF AND ShownRecords < PageSize
    End Function
'End EmailTemplate HasNextRow Function

End Class 'End EmailTemplate Class @2-A61BA892

Class clsEmailTemplateDataSource 'EmailTemplateDataSource Class @2-4E6450C4

'DataSource Variables @2-22C47CC3
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
    Public EmailTemplateID
    Public EmailTemplateID_param1
    Public EmailTemplateID_param2
    Public EmailTemplateName
    Public EmailTemplateToAddress
    Public EmailTemplateFromAddress
    Public EmailTemplateSubject
    Public EmailTemplateEmailType
    Public EmailTemplateUserLastUpdateBy
    Public EmailTemplateUserLastUpdateDateTime
    Public Alt_EmailTemplateID
    Public Alt_EmailTemplateID_param1
    Public Alt_EmailTemplateID_param2
    Public Alt_EmailTemplateName
    Public Alt_EmailTemplateToAddress
    Public Alt_EmailTemplateFromAddress
    Public Alt_EmailTemplateSubject
    Public Alt_EmailTemplateEmailType
    Public Alt_EmailTemplateUserLastUpdateBy
    Public Alt_EmailTemplateUserLastUpdateDateTime
'End DataSource Variables

'DataSource Class_Initialize Event @2-15963FA5
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set EmailTemplateID = CCCreateField("EmailTemplateID", "EmailTemplateID", ccsInteger, Empty, Recordset)
        Set EmailTemplateID_param1 = CCCreateField("EmailTemplateID_param1", "EmailTemplateID", ccsText, Empty, Recordset)
        Set EmailTemplateID_param2 = CCCreateField("EmailTemplateID_param2", "EmailTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Set EmailTemplateName = CCCreateField("EmailTemplateName", "EmailTemplateName", ccsText, Empty, Recordset)
        Set EmailTemplateToAddress = CCCreateField("EmailTemplateToAddress", "EmailTemplateToAddress", ccsText, Empty, Recordset)
        Set EmailTemplateFromAddress = CCCreateField("EmailTemplateFromAddress", "EmailTemplateFromAddress", ccsText, Empty, Recordset)
        Set EmailTemplateSubject = CCCreateField("EmailTemplateSubject", "EmailTemplateSubject", ccsText, Empty, Recordset)
        Set EmailTemplateEmailType = CCCreateField("EmailTemplateEmailType", "EmailTemplateEmailType", ccsText, Empty, Recordset)
        Set EmailTemplateUserLastUpdateBy = CCCreateField("EmailTemplateUserLastUpdateBy", "EmailTemplateUserLastUpdateBy", ccsText, Empty, Recordset)
        Set EmailTemplateUserLastUpdateDateTime = CCCreateField("EmailTemplateUserLastUpdateDateTime", "EmailTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateID = CCCreateField("Alt_EmailTemplateID", "EmailTemplateID", ccsInteger, Empty, Recordset)
        Set Alt_EmailTemplateID_param1 = CCCreateField("Alt_EmailTemplateID_param1", "EmailTemplateID", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateID_param2 = CCCreateField("Alt_EmailTemplateID_param2", "EmailTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateName = CCCreateField("Alt_EmailTemplateName", "EmailTemplateName", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateToAddress = CCCreateField("Alt_EmailTemplateToAddress", "EmailTemplateToAddress", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateFromAddress = CCCreateField("Alt_EmailTemplateFromAddress", "EmailTemplateFromAddress", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateSubject = CCCreateField("Alt_EmailTemplateSubject", "EmailTemplateSubject", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateEmailType = CCCreateField("Alt_EmailTemplateEmailType", "EmailTemplateEmailType", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateUserLastUpdateBy = CCCreateField("Alt_EmailTemplateUserLastUpdateBy", "EmailTemplateUserLastUpdateBy", ccsText, Empty, Recordset)
        Set Alt_EmailTemplateUserLastUpdateDateTime = CCCreateField("Alt_EmailTemplateUserLastUpdateDateTime", "EmailTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Fields.AddFields Array(EmailTemplateID,  EmailTemplateID_param1,  EmailTemplateID_param2,  EmailTemplateName,  EmailTemplateToAddress,  EmailTemplateFromAddress,  EmailTemplateSubject, _
             EmailTemplateEmailType,  EmailTemplateUserLastUpdateBy,  EmailTemplateUserLastUpdateDateTime,  Alt_EmailTemplateID,  Alt_EmailTemplateID_param1,  Alt_EmailTemplateID_param2,  Alt_EmailTemplateName,  Alt_EmailTemplateToAddress, _
             Alt_EmailTemplateFromAddress,  Alt_EmailTemplateSubject,  Alt_EmailTemplateEmailType,  Alt_EmailTemplateUserLastUpdateBy,  Alt_EmailTemplateUserLastUpdateDateTime)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing
        Orders = Array( _ 
            Array("Sorter_EmailID", "EmailTemplateID", ""), _
            Array("Sorter_EmailName", "EmailTemplateName", ""), _
            Array("Sorter_EmailToAddress", "EmailTemplateToAddress", ""), _
            Array("Sorter_EmailFromAddress", "EmailTemplateFromAddress", ""), _
            Array("Sorter_EmailSubject", "EmailTemplateSubject", ""))

        SQL = "SELECT TOP {SqlParam_endRecord}  *,  CONVERT(varchar, EmailTemplateUserLastUpdateDateTime, 20) AS EmailTemplateUserLastUpdateDateTime " & vbLf & _
        "FROM EmailTemplateArchive {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM EmailTemplateArchive"
        Where = ""
        Order = "EmailTemplateName, EmailTemplateUserLastUpdateDateTime desc"
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'SetOrder Method @2-68FC9576
    Sub SetOrder(Column, Direction)
        Order = Recordset.GetOrder(Order, Column, Direction, Orders)
    End Sub
'End SetOrder Method

'BuildTableWhere Method @2-0EE88FC4
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
            .AddParameter 2, "urls_keyword", ccsText, Empty, Empty, Empty, False
            .AddParameter 3, "urls_keyword", ccsText, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[EmailTemplateSiteID]", .getParamByID(1))
            .Criterion(2) = .Operation(opContains, False, "[EmailTemplateNickname]", .getParamByID(2))
            .Criterion(3) = .Operation(opContains, False, "[EmailTemplateName]", .getParamByID(3))
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

'Open Method @2-40984FC5
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

'DataSource Class_Terminate Event @2-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

End Class 'End EmailTemplateDataSource Class @2-A61BA892


%>
