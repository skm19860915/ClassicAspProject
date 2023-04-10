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

'Initialize Page @1-C981D840
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
Dim SendEmail
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminSendTestEmail.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminSendTestEmail.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Initialize Objects @1-349D03FF
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set SendEmail = new clsRecordSendEmail
Menu.Value = DHTMLMenu


CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-297237F1
SendEmail.Operation
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

'Show Page @1-7B32BE85
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, SendEmail))
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

'UnloadPage Sub @1-777D0CDA
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set SendEmail = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordSendEmail 'SendEmail Class @2-3E3744D2

'SendEmail Variables @2-BED57FD5

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
    Dim ToAddress
    Dim CCAddress
    Dim BCCAddress
    Dim FromAddress
    Dim EmailTemplateID
    Dim SiteID
'End SendEmail Variables

'SendEmail Class_Initialize Event @2-59EC3C0A
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
            FormSubmitted = (OperationMode(0) = "SendEmail")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "SendEmail"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set ToAddress = CCCreateControl(ccsTextBox, "ToAddress", Empty, ccsText, Empty, CCGetRequestParam("ToAddress", Method))
        Set CCAddress = CCCreateControl(ccsTextBox, "CCAddress", Empty, ccsText, Empty, CCGetRequestParam("CCAddress", Method))
        Set BCCAddress = CCCreateControl(ccsTextBox, "BCCAddress", Empty, ccsText, Empty, CCGetRequestParam("BCCAddress", Method))
        Set FromAddress = CCCreateControl(ccsTextBox, "FromAddress", Empty, ccsText, Empty, CCGetRequestParam("FromAddress", Method))
        Set EmailTemplateID = CCCreateList(ccsListBox, "EmailTemplateID", Empty, ccsText, CCGetRequestParam("EmailTemplateID", Method), Empty)
        EmailTemplateID.BoundColumn = "EmailTemplateNickname"
        EmailTemplateID.TextColumn = "EmailTemplateName"
        Set EmailTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT EmailTemplateName, EmailTemplateNickname  " & _
"FROM EmailTemplate {SQL_Where} {SQL_OrderBy}", "", "EmailTemplateName"))
        With EmailTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 1, "sesSiteID", ccsInteger, Empty, Empty, -1, False
            .Criterion(1) = .Operation(opEqual, False, "[EmailTemplateSiteID]", .getParamByID(1))
            .AssembledWhere = .Criterion(1)
        End With
        EmailTemplateID.DataSource.Where = EmailTemplateID.DataSource.WhereParameters.AssembledWhere
        EmailTemplateID.Required = True
        Set SiteID = CCCreateControl(ccsHidden, "SiteID", Empty, ccsInteger, Empty, CCGetRequestParam("SiteID", Method))
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(ToAddress, CCAddress, BCCAddress, FromAddress, EmailTemplateID, SiteID)
        If Not FormSubmitted Then
            If IsEmpty(SiteID.Value) Then _
                SiteID.Value = Session("SiteID")
        End If
    End Sub
'End SendEmail Class_Initialize Event

'SendEmail Class_Terminate Event @2-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End SendEmail Class_Terminate Event

'SendEmail Validate Method @2-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End SendEmail Validate Method

'SendEmail Operation Method @2-904D8152
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        Redirect = FileName & "?" & CCGetQueryString("QueryString", Array("ccsForm"))
    End Sub
'End SendEmail Operation Method

'SendEmail Show Method @2-739DFD49
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "SendEmail" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(ToAddress, CCAddress, BCCAddress, FromAddress, EmailTemplateID, SiteID))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors ToAddress.Errors
            Errors.AddErrors CCAddress.Errors
            Errors.AddErrors BCCAddress.Errors
            Errors.AddErrors FromAddress.Errors
            Errors.AddErrors EmailTemplateID.Errors
            Errors.AddErrors SiteID.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End SendEmail Show Method

End Class 'End SendEmail Class @2-A61BA892


%>
