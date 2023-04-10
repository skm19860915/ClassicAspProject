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

'Initialize Page @1-E5C418F0
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
Dim PageTemplate
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminPageTemplateArchiveView.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminPageTemplateArchiveView.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-3A73A510
CCSecurityRedirect "50", Empty
'End Authenticate User

'Initialize Objects @1-78207883
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set PageTemplate = new clsRecordPageTemplate
Menu.Value = DHTMLMenu

PageTemplate.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminPageTemplateArchiveView_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-DDD0B97A
PageTemplate.Operation
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

'Show Page @1-727D5048
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, PageTemplate))
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

'UnloadPage Sub @1-87B205FC
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set PageTemplate = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordPageTemplate 'PageTemplate Class @2-D79F4BEF

'PageTemplate Variables @2-B20EBE3A

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
    Dim Label1
    Dim PageTemplatePageType
    Dim PageTemplateParentPageTemplateID
    Dim PageTemplateStyleSheetPageTemplateID
    Dim PageTemplateBlockNames
    Dim SiteMainTemplateID
    Dim PageTemplateNickname
    Dim PageTemplateName
    Dim PageTemplateTitle
    Dim PageTemplateSiteID
    Dim PageTemplateUserLastUpdateBy
    Dim PageTemplateHead
    Dim PageTemplateContent
    Dim PageTemplateDynamicContent
    Dim PageTemplateExecuteASPFileName
    Dim UnrestrictedPages
    Dim RestrictedPages
    Dim PageTemplateRestrictedRedirectToPageTemplateID
    Dim PageTemplateRestrictedExceptions
    Dim PageTemplateRequiresLoginToAccess
    Dim PageTemplateRequiresLoginRedirectToPageTemplateID
    Dim PageTemplatePageTemplateSectionID
    Dim PageTemplateBlockList
    Dim Blocks
    Dim Button_Cancel
'End PageTemplate Variables

'PageTemplate Class_Initialize Event @2-70C9A474
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsPageTemplateDataSource
        Set Command = New clsCommand
        InsertAllowed = False
        UpdateAllowed = False
        DeleteAllowed = False
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "PageTemplate")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "PageTemplate"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set Label1 = CCCreateControl(ccsLabel, "Label1", Empty, ccsText, Empty, CCGetRequestParam("Label1", Method))
        Set PageTemplatePageType = CCCreateList(ccsListBox, "PageTemplatePageType", "Page Type", ccsText, CCGetRequestParam("PageTemplatePageType", Method), Empty)
        Set PageTemplatePageType.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("System", "User", "Template", "Block", "StyleSheet", "Process", "AJAX"), _
            Array("System", "User", "Template", "Block", "StyleSheet", "Process", "AJAX")))
        PageTemplatePageType.Required = True
        Set PageTemplateParentPageTemplateID = CCCreateList(ccsListBox, "PageTemplateParentPageTemplateID", Empty, ccsText, CCGetRequestParam("PageTemplateParentPageTemplateID", Method), Empty)
        PageTemplateParentPageTemplateID.BoundColumn = "PageTemplateNickname"
        PageTemplateParentPageTemplateID.TextColumn = "PageTemplateName"
        Set PageTemplateParentPageTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplate {SQL_Where} {SQL_OrderBy}", "", ""))
        With PageTemplateParentPageTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("expr56") = "Template"
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 1, "expr56", ccsText, Empty, Empty, Empty, False
            .AddParameter 2, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplatePageType]", .getParamByID(1))
            .Criterion(2) = .Operation(opEqual, False, "[PageTemplateSiteID]", .getParamByID(2))
            .AssembledWhere = .opAND(False, .Criterion(1), .Criterion(2))
        End With
        PageTemplateParentPageTemplateID.DataSource.Where = PageTemplateParentPageTemplateID.DataSource.WhereParameters.AssembledWhere
        Set PageTemplateStyleSheetPageTemplateID = CCCreateList(ccsListBox, "PageTemplateStyleSheetPageTemplateID", Empty, ccsText, CCGetRequestParam("PageTemplateStyleSheetPageTemplateID", Method), Empty)
        PageTemplateStyleSheetPageTemplateID.BoundColumn = "PageTemplateNickname"
        PageTemplateStyleSheetPageTemplateID.TextColumn = "PageTemplateName"
        Set PageTemplateStyleSheetPageTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplate {SQL_Where} {SQL_OrderBy}", "", ""))
        With PageTemplateStyleSheetPageTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("expr70") = "StyleSheet"
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 1, "expr70", ccsText, Empty, Empty, Empty, False
            .AddParameter 2, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplatePageType]", .getParamByID(1))
            .Criterion(2) = .Operation(opEqual, False, "[PageTemplateSiteID]", .getParamByID(2))
            .AssembledWhere = .opAND(False, .Criterion(1), .Criterion(2))
        End With
        PageTemplateStyleSheetPageTemplateID.DataSource.Where = PageTemplateStyleSheetPageTemplateID.DataSource.WhereParameters.AssembledWhere
        Set PageTemplateBlockNames = CCCreateControl(ccsTextBox, "PageTemplateBlockNames", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateBlockNames", Method))
        Set SiteMainTemplateID = CCCreateControl(ccsHidden, "SiteMainTemplateID", Empty, ccsText, Empty, CCGetRequestParam("SiteMainTemplateID", Method))
        Set PageTemplateNickname = CCCreateControl(ccsTextBox, "PageTemplateNickname", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateNickname", Method))
        PageTemplateNickname.Required = True
        Set PageTemplateName = CCCreateControl(ccsTextBox, "PageTemplateName", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateName", Method))
        PageTemplateName.Required = True
        Set PageTemplateTitle = CCCreateControl(ccsTextBox, "PageTemplateTitle", "Title", ccsText, Empty, CCGetRequestParam("PageTemplateTitle", Method))
        Set PageTemplateSiteID = CCCreateControl(ccsHidden, "PageTemplateSiteID", Empty, ccsInteger, Empty, CCGetRequestParam("PageTemplateSiteID", Method))
        Set PageTemplateUserLastUpdateBy = CCCreateControl(ccsHidden, "PageTemplateUserLastUpdateBy", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateUserLastUpdateBy", Method))
        Set PageTemplateHead = CCCreateControl(ccsTextArea, "PageTemplateHead", Empty, ccsMemo, Empty, CCGetRequestParam("PageTemplateHead", Method))
        Set PageTemplateContent = CCCreateControl(ccsTextArea, "PageTemplateContent", "Content", ccsMemo, Empty, CCGetRequestParam("PageTemplateContent", Method))
        Set PageTemplateDynamicContent = CCCreateControl(ccsTextArea, "PageTemplateDynamicContent", "Dynamic Content", ccsMemo, Empty, CCGetRequestParam("PageTemplateDynamicContent", Method))
        Set PageTemplateExecuteASPFileName = CCCreateControl(ccsTextBox, "PageTemplateExecuteASPFileName", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateExecuteASPFileName", Method))
        Set UnrestrictedPages = CCCreateList(ccsListBox, "UnrestrictedPages", Empty, ccsText, CCGetRequestMultipleParam("UnrestrictedPages", Method), Empty)
        UnrestrictedPages.BoundColumn = "PageTemplateNickname"
        UnrestrictedPages.TextColumn = "PageTemplateName"
        UnrestrictedPages.IsMultiple = True
        Set UnrestrictedPages.DataSource = CCCreateDataSource(dsSQL, DBSystem, "SELECT * " & _
"FROM PageTemplate " & _
"WHERE charIndex(',' + PageTemplateNickname + ',', ',' + (SELECT PageTemplateRestrictedExceptions FROM PageTemplate WHERE PageTemplateID = {PageTemplateID}) + ',') > 0")
        With UnrestrictedPages.DataSource
            .WhereParameters.AddParameter "PageTemplateID", "PageTemplateID", ccsText, Empty, Empty, -1, False
            .WhereParameters("PageTemplateID").Text = CCGetRequestParam("PageTemplateID", ccsGET)
        End With
        Set RestrictedPages = CCCreateList(ccsListBox, "RestrictedPages", Empty, ccsText, CCGetRequestMultipleParam("RestrictedPages", Method), Empty)
        RestrictedPages.BoundColumn = "PageTemplateNickname"
        RestrictedPages.TextColumn = "PageTemplateName"
        RestrictedPages.IsMultiple = True
        Set RestrictedPages.DataSource = CCCreateDataSource(dsSQL, DBSystem, "SELECT * " & _
"FROM PageTemplate " & _
"WHERE NOT charIndex(',' + PageTemplateNickname + ',', ISNULL(',' + (SELECT PageTemplateRestrictedExceptions FROM PageTemplate WHERE PageTemplateID = {PageTemplateID}) + ',', '')) > 0 " & _
"AND PageTemplatePageType IN ('System', 'User') " & _
"AND PageTemplateID <> {PageTemplateID} " & _
"AND PageTemplateSiteID = {SiteID} " & _
"ORDER BY PageTemplateName")
        With RestrictedPages.DataSource
            .WhereParameters.AddParameter "PageTemplateID", "PageTemplateID", ccsText, Empty, Empty, -1, False
            .WhereParameters.AddParameter "SiteID", "SiteID", ccsInteger, Empty, Empty, -1, False
            .WhereParameters("PageTemplateID").Text = CCGetRequestParam("PageTemplateID", ccsGET)
            .WhereParameters("SiteID").Text = Session("SiteID")
        End With
        Set PageTemplateRestrictedRedirectToPageTemplateID = CCCreateList(ccsListBox, "PageTemplateRestrictedRedirectToPageTemplateID", Empty, ccsText, CCGetRequestParam("PageTemplateRestrictedRedirectToPageTemplateID", Method), Empty)
        PageTemplateRestrictedRedirectToPageTemplateID.BoundColumn = "PageTemplateNickname"
        PageTemplateRestrictedRedirectToPageTemplateID.TextColumn = "PageTemplateName"
        Set PageTemplateRestrictedRedirectToPageTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplate {SQL_Where} {SQL_OrderBy}", "", "PageTemplateName"))
        With PageTemplateRestrictedRedirectToPageTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("urlPageTemplateID") = CCGetRequestParam("PageTemplateID", ccsGET)
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 2, "urlPageTemplateID", ccsInteger, Empty, Empty, Empty, False
            .AddParameter 3, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = "(PageTemplatePageType IN ('System', 'User'))"
            .Criterion(2) = .Operation(opNotEqual, False, "[PageTemplateID]", .getParamByID(2))
            .Criterion(3) = .Operation(opEqual, False, "[PageTemplateSiteID]", .getParamByID(3))
            .AssembledWhere = .opAND(False, .opAND(False, .Criterion(1), .Criterion(2)), .Criterion(3))
        End With
        PageTemplateRestrictedRedirectToPageTemplateID.DataSource.Where = PageTemplateRestrictedRedirectToPageTemplateID.DataSource.WhereParameters.AssembledWhere
        Set PageTemplateRestrictedExceptions = CCCreateControl(ccsHidden, "PageTemplateRestrictedExceptions", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateRestrictedExceptions", Method))
        Set PageTemplateRequiresLoginToAccess = CCCreateControl(ccsCheckBox, "PageTemplateRequiresLoginToAccess", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("PageTemplateRequiresLoginToAccess", Method))
        Set PageTemplateRequiresLoginRedirectToPageTemplateID = CCCreateList(ccsListBox, "PageTemplateRequiresLoginRedirectToPageTemplateID", Empty, ccsText, CCGetRequestParam("PageTemplateRequiresLoginRedirectToPageTemplateID", Method), Empty)
        PageTemplateRequiresLoginRedirectToPageTemplateID.BoundColumn = "PageTemplateNickname"
        PageTemplateRequiresLoginRedirectToPageTemplateID.TextColumn = "PageTemplateName"
        Set PageTemplateRequiresLoginRedirectToPageTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplate {SQL_Where} {SQL_OrderBy}", "", "PageTemplateName"))
        With PageTemplateRequiresLoginRedirectToPageTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("urlPageTemplateID") = CCGetRequestParam("PageTemplateID", ccsGET)
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 2, "urlPageTemplateID", ccsInteger, Empty, Empty, Empty, False
            .AddParameter 3, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = "(PageTemplatePageType IN ('System', 'User'))"
            .Criterion(2) = .Operation(opNotEqual, False, "[PageTemplateID]", .getParamByID(2))
            .Criterion(3) = .Operation(opEqual, False, "[PageTemplateSiteID]", .getParamByID(3))
            .AssembledWhere = .opAND(False, .opAND(False, .Criterion(1), .Criterion(2)), .Criterion(3))
        End With
        PageTemplateRequiresLoginRedirectToPageTemplateID.DataSource.Where = PageTemplateRequiresLoginRedirectToPageTemplateID.DataSource.WhereParameters.AssembledWhere
        Set PageTemplatePageTemplateSectionID = CCCreateList(ccsListBox, "PageTemplatePageTemplateSectionID", "Section", ccsText, CCGetRequestParam("PageTemplatePageTemplateSectionID", Method), Empty)
        PageTemplatePageTemplateSectionID.BoundColumn = "PageTemplateSectionNickname"
        PageTemplatePageTemplateSectionID.TextColumn = "PageTemplateSectionName"
        Set PageTemplatePageTemplateSectionID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplateSection {SQL_Where} {SQL_OrderBy}", "", ""))
        With PageTemplatePageTemplateSectionID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 1, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplateSectionSiteID]", .getParamByID(1))
            .AssembledWhere = .Criterion(1)
        End With
        PageTemplatePageTemplateSectionID.DataSource.Where = PageTemplatePageTemplateSectionID.DataSource.WhereParameters.AssembledWhere
        Set PageTemplateBlockList = CCCreateControl(ccsHidden, "PageTemplateBlockList", Empty, ccsText, Empty, CCGetRequestParam("PageTemplateBlockList", Method))
        Set Blocks = CCCreateControl(ccsLabel, "Blocks", Empty, ccsText, Empty, CCGetRequestParam("Blocks", Method))
        Blocks.HTML = True
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(PageTemplatePageType, PageTemplateParentPageTemplateID, PageTemplateStyleSheetPageTemplateID, PageTemplateBlockNames, SiteMainTemplateID, PageTemplateNickname, PageTemplateName, _
             PageTemplateTitle, PageTemplateSiteID, PageTemplateUserLastUpdateBy, PageTemplateHead, PageTemplateContent, PageTemplateDynamicContent, PageTemplateExecuteASPFileName, UnrestrictedPages, _
             RestrictedPages, PageTemplateRestrictedRedirectToPageTemplateID, PageTemplateRestrictedExceptions, PageTemplateRequiresLoginToAccess, PageTemplateRequiresLoginRedirectToPageTemplateID, PageTemplatePageTemplateSectionID, PageTemplateBlockList)
        If Not FormSubmitted Then
            If IsEmpty(PageTemplateSiteID.Value) Then _
                PageTemplateSiteID.Value = Session("SiteID")
        End If
    End Sub
'End PageTemplate Class_Initialize Event

'PageTemplate Initialize Method @2-9E3B52CF
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("urlPageTemplateID") = CCGetRequestParam("PageTemplateID", ccsGET)
            .Parameters("urlPageTemplateUserLastUpdateDateTime") = CCGetRequestParam("PageTemplateUserLastUpdateDateTime", ccsGET)
        End With
    End Sub
'End PageTemplate Initialize Method

'PageTemplate Class_Terminate Event @2-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End PageTemplate Class_Terminate Event

'PageTemplate Validate Method @2-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End PageTemplate Validate Method

'PageTemplate Operation Method @2-578EFC35
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_Cancel"
            If Button_Cancel.Pressed Then
                PressedButton = "Button_Cancel"
            End If
        End If
        Redirect = "AdminPageTemplateArchiveList.asp"
        If PressedButton = "Button_Cancel" Then
            If NOT Button_Cancel.OnClick Then
                Redirect = ""
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End PageTemplate Operation Method

'PageTemplate Show Method @2-C04DF918
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "PageTemplate" & IIf(EditMode, ":Edit", ""))
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
            Array(Label1,  PageTemplatePageType,  PageTemplateParentPageTemplateID,  PageTemplateStyleSheetPageTemplateID,  PageTemplateBlockNames,  SiteMainTemplateID,  PageTemplateNickname, _
                 PageTemplateName,  PageTemplateTitle,  PageTemplateSiteID,  PageTemplateUserLastUpdateBy,  PageTemplateHead,  PageTemplateContent,  PageTemplateDynamicContent,  PageTemplateExecuteASPFileName, _
                 UnrestrictedPages,  RestrictedPages,  PageTemplateRestrictedRedirectToPageTemplateID,  PageTemplateRestrictedExceptions,  PageTemplateRequiresLoginToAccess,  PageTemplateRequiresLoginRedirectToPageTemplateID,  PageTemplatePageTemplateSectionID,  PageTemplateBlockList,  Blocks,  Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        PageTemplatePageType.Value = Recordset.Fields("PageTemplatePageType")
                        PageTemplateParentPageTemplateID.Value = Recordset.Fields("PageTemplateParentPageTemplateID")
                        PageTemplateStyleSheetPageTemplateID.Value = Recordset.Fields("PageTemplateStyleSheetPageTemplateID")
                        PageTemplateBlockNames.Value = Recordset.Fields("PageTemplateBlockNames")
                        
                        PageTemplateNickname.Value = Recordset.Fields("PageTemplateNickname")
                        PageTemplateName.Value = Recordset.Fields("PageTemplateName")
                        PageTemplateTitle.Value = Recordset.Fields("PageTemplateTitle")
                        PageTemplateSiteID.Value = Recordset.Fields("PageTemplateSiteID")
                        PageTemplateUserLastUpdateBy.Value = Recordset.Fields("PageTemplateUserLastUpdateBy")
                        PageTemplateHead.Value = Recordset.Fields("PageTemplateHead")
                        PageTemplateContent.Value = Recordset.Fields("PageTemplateContent")
                        PageTemplateDynamicContent.Value = Recordset.Fields("PageTemplateDynamicContent")
                        PageTemplateExecuteASPFileName.Value = Recordset.Fields("PageTemplateExecuteASPFileName")
                        
                        
                        PageTemplateRestrictedRedirectToPageTemplateID.Value = Recordset.Fields("PageTemplateRestrictedRedirectToPageTemplateID")
                        PageTemplateRestrictedExceptions.Value = Recordset.Fields("PageTemplateRestrictedExceptions")
                        PageTemplateRequiresLoginToAccess.Value = Recordset.Fields("PageTemplateRequiresLoginToAccess")
                        PageTemplateRequiresLoginRedirectToPageTemplateID.Value = Recordset.Fields("PageTemplateRequiresLoginRedirectToPageTemplateID")
                        PageTemplatePageTemplateSectionID.Value = Recordset.Fields("PageTemplatePageTemplateSectionID")
                        PageTemplateBlockList.Value = Recordset.Fields("PageTemplateBlockList")
                    End If
                Else
                    EditMode = False
                End If
            End If
            If EditMode Then
                Label1.Value = Recordset.Fields("Label1")
                
            End If
        End If
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors PageTemplatePageType.Errors
            Errors.AddErrors PageTemplateParentPageTemplateID.Errors
            Errors.AddErrors PageTemplateStyleSheetPageTemplateID.Errors
            Errors.AddErrors PageTemplateBlockNames.Errors
            Errors.AddErrors SiteMainTemplateID.Errors
            Errors.AddErrors PageTemplateNickname.Errors
            Errors.AddErrors PageTemplateName.Errors
            Errors.AddErrors PageTemplateTitle.Errors
            Errors.AddErrors PageTemplateSiteID.Errors
            Errors.AddErrors PageTemplateUserLastUpdateBy.Errors
            Errors.AddErrors PageTemplateHead.Errors
            Errors.AddErrors PageTemplateContent.Errors
            Errors.AddErrors PageTemplateDynamicContent.Errors
            Errors.AddErrors PageTemplateExecuteASPFileName.Errors
            Errors.AddErrors UnrestrictedPages.Errors
            Errors.AddErrors RestrictedPages.Errors
            Errors.AddErrors PageTemplateRestrictedRedirectToPageTemplateID.Errors
            Errors.AddErrors PageTemplateRestrictedExceptions.Errors
            Errors.AddErrors PageTemplateRequiresLoginToAccess.Errors
            Errors.AddErrors PageTemplateRequiresLoginRedirectToPageTemplateID.Errors
            Errors.AddErrors PageTemplatePageTemplateSectionID.Errors
            Errors.AddErrors PageTemplateBlockList.Errors
            Errors.AddErrors DataSource.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End PageTemplate Show Method

End Class 'End PageTemplate Class @2-A61BA892

Class clsPageTemplateDataSource 'PageTemplateDataSource Class @2-FB673E35

'DataSource Variables @2-2B010187
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
    Public Label1
    Public PageTemplatePageType
    Public PageTemplateParentPageTemplateID
    Public PageTemplateStyleSheetPageTemplateID
    Public PageTemplateBlockNames
    Public PageTemplateNickname
    Public PageTemplateName
    Public PageTemplateTitle
    Public PageTemplateSiteID
    Public PageTemplateUserLastUpdateBy
    Public PageTemplateHead
    Public PageTemplateContent
    Public PageTemplateDynamicContent
    Public PageTemplateExecuteASPFileName
    Public PageTemplateRestrictedRedirectToPageTemplateID
    Public PageTemplateRestrictedExceptions
    Public PageTemplateRequiresLoginToAccess
    Public PageTemplateRequiresLoginRedirectToPageTemplateID
    Public PageTemplatePageTemplateSectionID
    Public PageTemplateBlockList
'End DataSource Variables

'DataSource Class_Initialize Event @2-32D3D3D0
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set Label1 = CCCreateField("Label1", "PageTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Set PageTemplatePageType = CCCreateField("PageTemplatePageType", "PageTemplatePageType", ccsText, Empty, Recordset)
        Set PageTemplateParentPageTemplateID = CCCreateField("PageTemplateParentPageTemplateID", "PageTemplateParentPageTemplateID", ccsText, Empty, Recordset)
        Set PageTemplateStyleSheetPageTemplateID = CCCreateField("PageTemplateStyleSheetPageTemplateID", "PageTemplateStyleSheetPageTemplateID", ccsText, Empty, Recordset)
        Set PageTemplateBlockNames = CCCreateField("PageTemplateBlockNames", "PageTemplateBlockNames", ccsText, Empty, Recordset)
        Set PageTemplateNickname = CCCreateField("PageTemplateNickname", "PageTemplateNickname", ccsText, Empty, Recordset)
        Set PageTemplateName = CCCreateField("PageTemplateName", "PageTemplateName", ccsText, Empty, Recordset)
        Set PageTemplateTitle = CCCreateField("PageTemplateTitle", "PageTemplateTitle", ccsText, Empty, Recordset)
        Set PageTemplateSiteID = CCCreateField("PageTemplateSiteID", "PageTemplateSiteID", ccsInteger, Empty, Recordset)
        Set PageTemplateUserLastUpdateBy = CCCreateField("PageTemplateUserLastUpdateBy", "PageTemplateUserLastUpdateBy", ccsText, Empty, Recordset)
        Set PageTemplateHead = CCCreateField("PageTemplateHead", "PageTemplateHead", ccsMemo, Empty, Recordset)
        Set PageTemplateContent = CCCreateField("PageTemplateContent", "PageTemplateContent", ccsMemo, Empty, Recordset)
        Set PageTemplateDynamicContent = CCCreateField("PageTemplateDynamicContent", "PageTemplateDynamicContent", ccsMemo, Empty, Recordset)
        Set PageTemplateExecuteASPFileName = CCCreateField("PageTemplateExecuteASPFileName", "PageTemplateExecuteASPFileName", ccsText, Empty, Recordset)
        Set PageTemplateRestrictedRedirectToPageTemplateID = CCCreateField("PageTemplateRestrictedRedirectToPageTemplateID", "PageTemplateRestrictedRedirectToPageTemplateID", ccsText, Empty, Recordset)
        Set PageTemplateRestrictedExceptions = CCCreateField("PageTemplateRestrictedExceptions", "PageTemplateRestrictedExceptions", ccsText, Empty, Recordset)
        Set PageTemplateRequiresLoginToAccess = CCCreateField("PageTemplateRequiresLoginToAccess", "PageTemplateRequiresLoginToAccess", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set PageTemplateRequiresLoginRedirectToPageTemplateID = CCCreateField("PageTemplateRequiresLoginRedirectToPageTemplateID", "PageTemplateRequiresLoginRedirectToPageTemplateID", ccsText, Empty, Recordset)
        Set PageTemplatePageTemplateSectionID = CCCreateField("PageTemplatePageTemplateSectionID", "PageTemplatePageTemplateSectionID", ccsText, Empty, Recordset)
        Set PageTemplateBlockList = CCCreateField("PageTemplateBlockList", "PageTemplateBlockList", ccsText, Empty, Recordset)
        Fields.AddFields Array(Label1,  PageTemplatePageType,  PageTemplateParentPageTemplateID,  PageTemplateStyleSheetPageTemplateID,  PageTemplateBlockNames,  PageTemplateNickname,  PageTemplateName, _
             PageTemplateTitle,  PageTemplateSiteID,  PageTemplateUserLastUpdateBy,  PageTemplateHead,  PageTemplateContent,  PageTemplateDynamicContent,  PageTemplateExecuteASPFileName,  PageTemplateRestrictedRedirectToPageTemplateID, _
             PageTemplateRestrictedExceptions,  PageTemplateRequiresLoginToAccess,  PageTemplateRequiresLoginRedirectToPageTemplateID,  PageTemplatePageTemplateSectionID,  PageTemplateBlockList)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1 *  " & vbLf & _
        "FROM PageTemplateArchive " & vbLf & _
        "WHERE PageTemplateID = {PageTemplateID} " & vbLf & _
        "AND CONVERT(varchar, PageTemplateUserLastUpdateDateTime, 20) = '{PageTemplateUserLastUpdateDateTime}'"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @2-7B5B770B
    Public Sub BuildTableWhere()
        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter "PageTemplateID", "urlPageTemplateID", ccsInteger, Empty, Empty, 0, False
            .AddParameter "PageTemplateUserLastUpdateDateTime", "urlPageTemplateUserLastUpdateDateTime", ccsText, Empty, Empty, Empty, False
            AllParamsSet = .AllParamsSet
        End With
    End Sub
'End BuildTableWhere Method

'Open Method @2-251BBF69
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
        BuildTableWhere
        Set Cmd.WhereParameters = WhereParameters
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

'DataSource Class_Terminate Event @2-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

End Class 'End PageTemplateDataSource Class @2-A61BA892


%>
