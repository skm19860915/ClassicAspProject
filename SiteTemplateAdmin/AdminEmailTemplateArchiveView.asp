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

'Initialize Page @1-834F84B7
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
Dim EmailTemplate
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminEmailTemplateArchiveView.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminEmailTemplateArchiveView.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-21F7A3FF
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set EmailTemplate = new clsRecordEmailTemplate
Menu.Value = DHTMLMenu

EmailTemplate.Initialize DBSystem

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-8DA5A151
EmailTemplate.Operation
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

'Show Page @1-B65B2E6C
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, EmailTemplate))
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

'UnloadPage Sub @1-094737CD
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set EmailTemplate = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordEmailTemplate 'EmailTemplate Class @2-CF74E6A3

'EmailTemplate Variables @2-023CF750

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
    Dim EmailTemplateEmailType
    Dim EmailTemplateParentEmailTemplateID
    Dim EmailTemplateSection
    Dim EmailTemplateNickname
    Dim EmailTemplateName
    Dim EmailTemplateSiteID
    Dim EmailTemplateUserLastUpdateBy
    Dim EmailTemplateToAddress
    Dim EmailTemplateFromAddress
    Dim EmailTemplateFromName
    Dim EmailTemplateReplyToAddress
    Dim EmailTemplateReplyToName
    Dim EmailTemplateBounceBackEmail
    Dim EmailTemplateEmbedImages
    Dim EmailTemplateImageURL
    Dim EmailTemplateImagePath
    Dim EmailTemplateImagePhysicalPath
    Dim EmailTemplateSubject
    Dim EmailTemplateBody
    Dim EmailTemplateBodyTextOnly
    Dim EmailTemplateBodyDynamicContent
    Dim Button_Cancel
'End EmailTemplate Variables

'EmailTemplate Class_Initialize Event @2-A11092C1
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsEmailTemplateDataSource
        Set Command = New clsCommand
        InsertAllowed = False
        UpdateAllowed = False
        DeleteAllowed = False
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "EmailTemplate")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "EmailTemplate"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set Label1 = CCCreateControl(ccsLabel, "Label1", Empty, ccsText, Empty, CCGetRequestParam("Label1", Method))
        Set EmailTemplateEmailType = CCCreateList(ccsListBox, "EmailTemplateEmailType", "Email Type", ccsText, CCGetRequestParam("EmailTemplateEmailType", Method), Empty)
        Set EmailTemplateEmailType.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("System", "User", "Template"), _
            Array("System", "User", "Template")))
        EmailTemplateEmailType.Required = True
        Set EmailTemplateParentEmailTemplateID = CCCreateList(ccsListBox, "EmailTemplateParentEmailTemplateID", Empty, ccsText, CCGetRequestParam("EmailTemplateParentEmailTemplateID", Method), Empty)
        EmailTemplateParentEmailTemplateID.BoundColumn = "EmailTemplateNickname"
        EmailTemplateParentEmailTemplateID.TextColumn = "EmailTemplateName"
        Set EmailTemplateParentEmailTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM EmailTemplate {SQL_Where} {SQL_OrderBy}", "", ""))
        With EmailTemplateParentEmailTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("expr89") = "Template"
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 1, "expr89", ccsText, Empty, Empty, Empty, False
            .AddParameter 2, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[EmailTemplateEmailType]", .getParamByID(1))
            .Criterion(2) = .Operation(opEqual, False, "[EmailTemplateSiteID]", .getParamByID(2))
            .AssembledWhere = .opAND(False, .Criterion(1), .Criterion(2))
        End With
        EmailTemplateParentEmailTemplateID.DataSource.Where = EmailTemplateParentEmailTemplateID.DataSource.WhereParameters.AssembledWhere
        Set EmailTemplateSection = CCCreateList(ccsListBox, "EmailTemplateSection", "Section", ccsText, CCGetRequestParam("EmailTemplateSection", Method), Empty)
        Set EmailTemplateSection.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("0", "1", "2", "3"), _
            Array("Client Confirmation", "Client Notification", "Admin Confirmation", "Admin Notification")))
        Set EmailTemplateNickname = CCCreateControl(ccsTextBox, "EmailTemplateNickname", "Nickname", ccsText, Empty, CCGetRequestParam("EmailTemplateNickname", Method))
        EmailTemplateNickname.Required = True
        Set EmailTemplateName = CCCreateControl(ccsTextBox, "EmailTemplateName", "Name", ccsText, Empty, CCGetRequestParam("EmailTemplateName", Method))
        Set EmailTemplateSiteID = CCCreateControl(ccsHidden, "EmailTemplateSiteID", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateSiteID", Method))
        Set EmailTemplateUserLastUpdateBy = CCCreateControl(ccsHidden, "EmailTemplateUserLastUpdateBy", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateUserLastUpdateBy", Method))
        Set EmailTemplateToAddress = CCCreateControl(ccsTextBox, "EmailTemplateToAddress", "To Address", ccsText, Empty, CCGetRequestParam("EmailTemplateToAddress", Method))
        Set EmailTemplateFromAddress = CCCreateControl(ccsTextBox, "EmailTemplateFromAddress", "From Address", ccsText, Empty, CCGetRequestParam("EmailTemplateFromAddress", Method))
        Set EmailTemplateFromName = CCCreateControl(ccsTextBox, "EmailTemplateFromName", "From Name", ccsText, Empty, CCGetRequestParam("EmailTemplateFromName", Method))
        Set EmailTemplateReplyToAddress = CCCreateControl(ccsTextBox, "EmailTemplateReplyToAddress", "Reply To Address", ccsText, Empty, CCGetRequestParam("EmailTemplateReplyToAddress", Method))
        Set EmailTemplateReplyToName = CCCreateControl(ccsTextBox, "EmailTemplateReplyToName", "Reply To Name", ccsText, Empty, CCGetRequestParam("EmailTemplateReplyToName", Method))
        Set EmailTemplateBounceBackEmail = CCCreateControl(ccsTextBox, "EmailTemplateBounceBackEmail", "Bounce Back Email", ccsText, Empty, CCGetRequestParam("EmailTemplateBounceBackEmail", Method))
        Set EmailTemplateEmbedImages = CCCreateControl(ccsCheckBox, "EmailTemplateEmbedImages", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("EmailTemplateEmbedImages", Method))
        Set EmailTemplateImageURL = CCCreateControl(ccsTextBox, "EmailTemplateImageURL", "Image URL", ccsMemo, Empty, CCGetRequestParam("EmailTemplateImageURL", Method))
        Set EmailTemplateImagePath = CCCreateControl(ccsTextBox, "EmailTemplateImagePath", "Image Path", ccsText, Empty, CCGetRequestParam("EmailTemplateImagePath", Method))
        Set EmailTemplateImagePhysicalPath = CCCreateControl(ccsHidden, "EmailTemplateImagePhysicalPath", Empty, ccsText, Empty, CCGetRequestParam("EmailTemplateImagePhysicalPath", Method))
        Set EmailTemplateSubject = CCCreateControl(ccsTextBox, "EmailTemplateSubject", "Subject", ccsText, Empty, CCGetRequestParam("EmailTemplateSubject", Method))
        Set EmailTemplateBody = CCCreateControl(ccsTextArea, "EmailTemplateBody", "Body", ccsMemo, Empty, CCGetRequestParam("EmailTemplateBody", Method))
        Set EmailTemplateBodyTextOnly = CCCreateControl(ccsTextArea, "EmailTemplateBodyTextOnly", "Body Text Only", ccsMemo, Empty, CCGetRequestParam("EmailTemplateBodyTextOnly", Method))
        Set EmailTemplateBodyDynamicContent = CCCreateControl(ccsTextArea, "EmailTemplateBodyDynamicContent", "Body Dynamic Content", ccsMemo, Empty, CCGetRequestParam("EmailTemplateBodyDynamicContent", Method))
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(EmailTemplateEmailType, EmailTemplateParentEmailTemplateID, EmailTemplateSection, EmailTemplateNickname, EmailTemplateName, EmailTemplateSiteID, EmailTemplateUserLastUpdateBy, _
             EmailTemplateToAddress, EmailTemplateFromAddress, EmailTemplateFromName, EmailTemplateReplyToAddress, EmailTemplateReplyToName, EmailTemplateBounceBackEmail, EmailTemplateEmbedImages, EmailTemplateImageURL, _
             EmailTemplateImagePath, EmailTemplateImagePhysicalPath, EmailTemplateSubject, EmailTemplateBody, EmailTemplateBodyTextOnly, EmailTemplateBodyDynamicContent)
        If Not FormSubmitted Then
            If IsEmpty(EmailTemplateSiteID.Value) Then _
                EmailTemplateSiteID.Value = Session("SiteID")
        End If
    End Sub
'End EmailTemplate Class_Initialize Event

'EmailTemplate Initialize Method @2-D16D57FB
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("urlEmailTemplateID") = CCGetRequestParam("EmailTemplateID", ccsGET)
            .Parameters("urlEmailTemplateUserLastUpdateDateTime") = CCGetRequestParam("EmailTemplateUserLastUpdateDateTime", ccsGET)
        End With
    End Sub
'End EmailTemplate Initialize Method

'EmailTemplate Class_Terminate Event @2-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End EmailTemplate Class_Terminate Event

'EmailTemplate Validate Method @2-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End EmailTemplate Validate Method

'EmailTemplate Operation Method @2-333BFAA9
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_Cancel"
            If Button_Cancel.Pressed Then
                PressedButton = "Button_Cancel"
            End If
        End If
        Redirect = "AdminEmailTemplateArchiveList.asp"
        If PressedButton = "Button_Cancel" Then
            If NOT Button_Cancel.OnClick Then
                Redirect = ""
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End EmailTemplate Operation Method

'EmailTemplate Show Method @2-411E5E2D
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "EmailTemplate" & IIf(EditMode, ":Edit", ""))
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
            Array(Label1,  EmailTemplateEmailType,  EmailTemplateParentEmailTemplateID,  EmailTemplateSection,  EmailTemplateNickname,  EmailTemplateName,  EmailTemplateSiteID, _
                 EmailTemplateUserLastUpdateBy,  EmailTemplateToAddress,  EmailTemplateFromAddress,  EmailTemplateFromName,  EmailTemplateReplyToAddress,  EmailTemplateReplyToName,  EmailTemplateBounceBackEmail,  EmailTemplateEmbedImages, _
                 EmailTemplateImageURL,  EmailTemplateImagePath,  EmailTemplateImagePhysicalPath,  EmailTemplateSubject,  EmailTemplateBody,  EmailTemplateBodyTextOnly,  EmailTemplateBodyDynamicContent,  Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        EmailTemplateEmailType.Value = Recordset.Fields("EmailTemplateEmailType")
                        EmailTemplateParentEmailTemplateID.Value = Recordset.Fields("EmailTemplateParentEmailTemplateID")
                        EmailTemplateSection.Value = Recordset.Fields("EmailTemplateSection")
                        EmailTemplateNickname.Value = Recordset.Fields("EmailTemplateNickname")
                        EmailTemplateName.Value = Recordset.Fields("EmailTemplateName")
                        EmailTemplateSiteID.Value = Recordset.Fields("EmailTemplateSiteID")
                        EmailTemplateUserLastUpdateBy.Value = Recordset.Fields("EmailTemplateUserLastUpdateBy")
                        EmailTemplateToAddress.Value = Recordset.Fields("EmailTemplateToAddress")
                        EmailTemplateFromAddress.Value = Recordset.Fields("EmailTemplateFromAddress")
                        EmailTemplateFromName.Value = Recordset.Fields("EmailTemplateFromName")
                        EmailTemplateReplyToAddress.Value = Recordset.Fields("EmailTemplateReplyToAddress")
                        EmailTemplateReplyToName.Value = Recordset.Fields("EmailTemplateReplyToName")
                        EmailTemplateBounceBackEmail.Value = Recordset.Fields("EmailTemplateBounceBackEmail")
                        EmailTemplateEmbedImages.Value = Recordset.Fields("EmailTemplateEmbedImages")
                        EmailTemplateImageURL.Value = Recordset.Fields("EmailTemplateImageURL")
                        EmailTemplateImagePath.Value = Recordset.Fields("EmailTemplateImagePath")
                        EmailTemplateImagePhysicalPath.Value = Recordset.Fields("EmailTemplateImagePhysicalPath")
                        EmailTemplateSubject.Value = Recordset.Fields("EmailTemplateSubject")
                        EmailTemplateBody.Value = Recordset.Fields("EmailTemplateBody")
                        EmailTemplateBodyTextOnly.Value = Recordset.Fields("EmailTemplateBodyTextOnly")
                        EmailTemplateBodyDynamicContent.Value = Recordset.Fields("EmailTemplateBodyDynamicContent")
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
            Errors.AddErrors EmailTemplateEmailType.Errors
            Errors.AddErrors EmailTemplateParentEmailTemplateID.Errors
            Errors.AddErrors EmailTemplateSection.Errors
            Errors.AddErrors EmailTemplateNickname.Errors
            Errors.AddErrors EmailTemplateName.Errors
            Errors.AddErrors EmailTemplateSiteID.Errors
            Errors.AddErrors EmailTemplateUserLastUpdateBy.Errors
            Errors.AddErrors EmailTemplateToAddress.Errors
            Errors.AddErrors EmailTemplateFromAddress.Errors
            Errors.AddErrors EmailTemplateFromName.Errors
            Errors.AddErrors EmailTemplateReplyToAddress.Errors
            Errors.AddErrors EmailTemplateReplyToName.Errors
            Errors.AddErrors EmailTemplateBounceBackEmail.Errors
            Errors.AddErrors EmailTemplateEmbedImages.Errors
            Errors.AddErrors EmailTemplateImageURL.Errors
            Errors.AddErrors EmailTemplateImagePath.Errors
            Errors.AddErrors EmailTemplateImagePhysicalPath.Errors
            Errors.AddErrors EmailTemplateSubject.Errors
            Errors.AddErrors EmailTemplateBody.Errors
            Errors.AddErrors EmailTemplateBodyTextOnly.Errors
            Errors.AddErrors EmailTemplateBodyDynamicContent.Errors
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
'End EmailTemplate Show Method

End Class 'End EmailTemplate Class @2-A61BA892

Class clsEmailTemplateDataSource 'EmailTemplateDataSource Class @2-4E6450C4

'DataSource Variables @2-DE99D5B7
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
    Public EmailTemplateEmailType
    Public EmailTemplateParentEmailTemplateID
    Public EmailTemplateSection
    Public EmailTemplateNickname
    Public EmailTemplateName
    Public EmailTemplateSiteID
    Public EmailTemplateUserLastUpdateBy
    Public EmailTemplateToAddress
    Public EmailTemplateFromAddress
    Public EmailTemplateFromName
    Public EmailTemplateReplyToAddress
    Public EmailTemplateReplyToName
    Public EmailTemplateBounceBackEmail
    Public EmailTemplateEmbedImages
    Public EmailTemplateImageURL
    Public EmailTemplateImagePath
    Public EmailTemplateImagePhysicalPath
    Public EmailTemplateSubject
    Public EmailTemplateBody
    Public EmailTemplateBodyTextOnly
    Public EmailTemplateBodyDynamicContent
'End DataSource Variables

'DataSource Class_Initialize Event @2-7D6A5C11
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set Label1 = CCCreateField("Label1", "EmailTemplateUserLastUpdateDateTime", ccsText, Empty, Recordset)
        Set EmailTemplateEmailType = CCCreateField("EmailTemplateEmailType", "EmailTemplateEmailType", ccsText, Empty, Recordset)
        Set EmailTemplateParentEmailTemplateID = CCCreateField("EmailTemplateParentEmailTemplateID", "EmailTemplateParentEmailTemplateID", ccsText, Empty, Recordset)
        Set EmailTemplateSection = CCCreateField("EmailTemplateSection", "EmailTemplateSection", ccsText, Empty, Recordset)
        Set EmailTemplateNickname = CCCreateField("EmailTemplateNickname", "EmailTemplateNickname", ccsText, Empty, Recordset)
        Set EmailTemplateName = CCCreateField("EmailTemplateName", "EmailTemplateName", ccsText, Empty, Recordset)
        Set EmailTemplateSiteID = CCCreateField("EmailTemplateSiteID", "EmailTemplateSiteID", ccsText, Empty, Recordset)
        Set EmailTemplateUserLastUpdateBy = CCCreateField("EmailTemplateUserLastUpdateBy", "EmailTemplateUserLastUpdateBy", ccsText, Empty, Recordset)
        Set EmailTemplateToAddress = CCCreateField("EmailTemplateToAddress", "EmailTemplateToAddress", ccsText, Empty, Recordset)
        Set EmailTemplateFromAddress = CCCreateField("EmailTemplateFromAddress", "EmailTemplateFromAddress", ccsText, Empty, Recordset)
        Set EmailTemplateFromName = CCCreateField("EmailTemplateFromName", "EmailTemplateFromName", ccsText, Empty, Recordset)
        Set EmailTemplateReplyToAddress = CCCreateField("EmailTemplateReplyToAddress", "EmailTemplateReplyToAddress", ccsText, Empty, Recordset)
        Set EmailTemplateReplyToName = CCCreateField("EmailTemplateReplyToName", "EmailTemplateReplyToName", ccsText, Empty, Recordset)
        Set EmailTemplateBounceBackEmail = CCCreateField("EmailTemplateBounceBackEmail", "EmailTemplateBounceBackEmail", ccsText, Empty, Recordset)
        Set EmailTemplateEmbedImages = CCCreateField("EmailTemplateEmbedImages", "EmailTemplateEmbedImages", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set EmailTemplateImageURL = CCCreateField("EmailTemplateImageURL", "EmailTemplateImageURL", ccsMemo, Empty, Recordset)
        Set EmailTemplateImagePath = CCCreateField("EmailTemplateImagePath", "EmailTemplateImagePath", ccsText, Empty, Recordset)
        Set EmailTemplateImagePhysicalPath = CCCreateField("EmailTemplateImagePhysicalPath", "EmailTemplateImagePhysicalPath", ccsText, Empty, Recordset)
        Set EmailTemplateSubject = CCCreateField("EmailTemplateSubject", "EmailTemplateSubject", ccsText, Empty, Recordset)
        Set EmailTemplateBody = CCCreateField("EmailTemplateBody", "EmailTemplateBody", ccsMemo, Empty, Recordset)
        Set EmailTemplateBodyTextOnly = CCCreateField("EmailTemplateBodyTextOnly", "EmailTemplateBodyTextOnly", ccsMemo, Empty, Recordset)
        Set EmailTemplateBodyDynamicContent = CCCreateField("EmailTemplateBodyDynamicContent", "EmailTemplateBodyDynamicContent", ccsMemo, Empty, Recordset)
        Fields.AddFields Array(Label1,  EmailTemplateEmailType,  EmailTemplateParentEmailTemplateID,  EmailTemplateSection,  EmailTemplateNickname,  EmailTemplateName,  EmailTemplateSiteID, _
             EmailTemplateUserLastUpdateBy,  EmailTemplateToAddress,  EmailTemplateFromAddress,  EmailTemplateFromName,  EmailTemplateReplyToAddress,  EmailTemplateReplyToName,  EmailTemplateBounceBackEmail,  EmailTemplateEmbedImages, _
             EmailTemplateImageURL,  EmailTemplateImagePath,  EmailTemplateImagePhysicalPath,  EmailTemplateSubject,  EmailTemplateBody,  EmailTemplateBodyTextOnly,  EmailTemplateBodyDynamicContent)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1 *  " & vbLf & _
        "FROM EmailTemplateArchive " & vbLf & _
        "WHERE EmailTemplateID = {EmailTemplateID} " & vbLf & _
        "AND CONVERT(varchar, EmailTemplateUserLastUpdateDateTime, 20) = '{EmailTemplateUserLastUpdateDateTime}'"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @2-5F4C9093
    Public Sub BuildTableWhere()
        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter "EmailTemplateID", "urlEmailTemplateID", ccsInteger, Empty, Empty, 0, False
            .AddParameter "EmailTemplateUserLastUpdateDateTime", "urlEmailTemplateUserLastUpdateDateTime", ccsText, Empty, Empty, Empty, False
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

End Class 'End EmailTemplateDataSource Class @2-A61BA892


%>
