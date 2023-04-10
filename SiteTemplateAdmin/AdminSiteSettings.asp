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

'Initialize Page @1-59BB059F
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
Dim Site
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminSiteSettings.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminSiteSettings.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-3A73A510
CCSecurityRedirect "50", Empty
'End Authenticate User

'Initialize Objects @1-77130B85
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set Site = new clsRecordSite
Menu.Value = DHTMLMenu

Site.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminSiteSettings_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-2995C4C5
Site.Operation
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

'Show Page @1-1DFF07B2
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, Site))
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

'UnloadPage Sub @1-A9F56734
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set Site = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordSite 'Site Class @55-74A8C9ED

'Site Variables @55-4A2D9C14

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
    Dim SiteName
    Dim SiteDomainName
    Dim SiteDescription
    Dim SiteMaxTemplateBlocks
    Dim SiteMainTemplateID
    Dim SiteURLEncryptionKey
    Dim SiteDataEncryptionKey
    Dim SiteTrackingString
    Dim SiteTrackingLinkLength
    Dim SiteTrackingCookieTTL
    Dim SiteDefaultURL
    Dim SiteRefererCookieTTL
    Dim SiteDefaultPageTemplateID
    Dim SiteAuthorization
    Dim SiteAuthorizeLogin
    Dim SiteAuthorizeTransactionKey
    Dim SiteAuthorizeDescription
    Dim SiteProPrice
    Dim SiteUseEditor
    Dim SiteDebugMode
    Dim SiteDebugIP
    Dim SiteDebugModeCookieTTL
    Dim DebugDomainName
    Dim SiteContentManagementIP
    Dim SiteContentManagementModeCookieTTL
    Dim SiteRedirectPageOnError
    Dim SiteLoginPage
    Dim SiteUploadPath
    Dim SitePageTemplateExecuteASPPath
    Dim SiteCustom404ScriptName
    Dim Button_Insert
    Dim Button_Update
    Dim Button_Cancel
'End Site Variables

'Site Class_Initialize Event @55-A534175E
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsSiteDataSource
        Set Command = New clsCommand
        InsertAllowed = True
        UpdateAllowed = True
        DeleteAllowed = False
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "Site")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "Site"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set SiteName = CCCreateControl(ccsTextBox, "SiteName", "Name", ccsText, Empty, CCGetRequestParam("SiteName", Method))
        Set SiteDomainName = CCCreateControl(ccsTextBox, "SiteDomainName", "Domain Name", ccsText, Empty, CCGetRequestParam("SiteDomainName", Method))
        Set SiteDescription = CCCreateControl(ccsTextArea, "SiteDescription", "Description", ccsMemo, Empty, CCGetRequestParam("SiteDescription", Method))
        Set SiteMaxTemplateBlocks = CCCreateControl(ccsTextBox, "SiteMaxTemplateBlocks", "Max Template Blocks", ccsInteger, Empty, CCGetRequestParam("SiteMaxTemplateBlocks", Method))
        Set SiteMainTemplateID = CCCreateList(ccsListBox, "SiteMainTemplateID", "Main Template", ccsText, CCGetRequestParam("SiteMainTemplateID", Method), Empty)
        SiteMainTemplateID.BoundColumn = "PageTemplateNickname"
        SiteMainTemplateID.TextColumn = "PageTemplateName"
        Set SiteMainTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplate {SQL_Where} {SQL_OrderBy}", "", ""))
        With SiteMainTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("sesSiteID") = Session("SiteID")
            .ParameterSources("expr85") = "Template"
            .AddParameter 1, "sesSiteID", ccsInteger, Empty, Empty, -1, False
            .AddParameter 2, "expr85", ccsText, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplateSiteID]", .getParamByID(1))
            .Criterion(2) = .Operation(opEqual, False, "[PageTemplatePageType]", .getParamByID(2))
            .AssembledWhere = .opAND(False, .Criterion(1), .Criterion(2))
        End With
        SiteMainTemplateID.DataSource.Where = SiteMainTemplateID.DataSource.WhereParameters.AssembledWhere
        Set SiteURLEncryptionKey = CCCreateControl(ccsTextBox, "SiteURLEncryptionKey", "URL Encryption Key", ccsText, Empty, CCGetRequestParam("SiteURLEncryptionKey", Method))
        Set SiteDataEncryptionKey = CCCreateControl(ccsTextBox, "SiteDataEncryptionKey", "Data Encryption Key", ccsText, Empty, CCGetRequestParam("SiteDataEncryptionKey", Method))
        Set SiteTrackingString = CCCreateControl(ccsTextBox, "SiteTrackingString", "Tracking String", ccsText, Empty, CCGetRequestParam("SiteTrackingString", Method))
        Set SiteTrackingLinkLength = CCCreateControl(ccsTextBox, "SiteTrackingLinkLength", "Tracking Link Length", ccsInteger, Empty, CCGetRequestParam("SiteTrackingLinkLength", Method))
        Set SiteTrackingCookieTTL = CCCreateControl(ccsTextBox, "SiteTrackingCookieTTL", "Tracking Cookie TTL", ccsInteger, Empty, CCGetRequestParam("SiteTrackingCookieTTL", Method))
        Set SiteDefaultURL = CCCreateControl(ccsTextBox, "SiteDefaultURL", "Default URL", ccsText, Empty, CCGetRequestParam("SiteDefaultURL", Method))
        Set SiteRefererCookieTTL = CCCreateControl(ccsTextBox, "SiteRefererCookieTTL", "Referer Cookie TTL", ccsInteger, Empty, CCGetRequestParam("SiteRefererCookieTTL", Method))
        Set SiteDefaultPageTemplateID = CCCreateList(ccsListBox, "SiteDefaultPageTemplateID", "Default Page Template ID", ccsText, CCGetRequestParam("SiteDefaultPageTemplateID", Method), Empty)
        SiteDefaultPageTemplateID.BoundColumn = "PageTemplateNickname"
        SiteDefaultPageTemplateID.TextColumn = "PageTemplateName"
        Set SiteDefaultPageTemplateID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM PageTemplate {SQL_Where} {SQL_OrderBy}", "", ""))
        With SiteDefaultPageTemplateID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("sesSiteID") = Session("SiteID")
            .ParameterSources("expr91") = "System"
            .ParameterSources("expr92") = "User"
            .AddParameter 1, "sesSiteID", ccsInteger, Empty, Empty, -1, False
            .AddParameter 2, "expr91", ccsText, Empty, Empty, Empty, False
            .AddParameter 3, "expr92", ccsText, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplateSiteID]", .getParamByID(1))
            .Criterion(2) = .Operation(opEqual, False, "[PageTemplatePageType]", .getParamByID(2))
            .Criterion(3) = .Operation(opEqual, False, "[PageTemplatePageType]", .getParamByID(3))
            .AssembledWhere = .opAND(False, .Criterion(1), .opOR(True, .Criterion(2), .Criterion(3)))
        End With
        SiteDefaultPageTemplateID.DataSource.Where = SiteDefaultPageTemplateID.DataSource.WhereParameters.AssembledWhere
        Set SiteAuthorization = CCCreateControl(ccsTextBox, "SiteAuthorization", "Authorization", ccsText, Empty, CCGetRequestParam("SiteAuthorization", Method))
        Set SiteAuthorizeLogin = CCCreateControl(ccsTextBox, "SiteAuthorizeLogin", "Authorize Login", ccsText, Empty, CCGetRequestParam("SiteAuthorizeLogin", Method))
        Set SiteAuthorizeTransactionKey = CCCreateControl(ccsTextBox, "SiteAuthorizeTransactionKey", "Authorize Transaction Key", ccsText, Empty, CCGetRequestParam("SiteAuthorizeTransactionKey", Method))
        Set SiteAuthorizeDescription = CCCreateControl(ccsTextBox, "SiteAuthorizeDescription", "Authorize Description", ccsText, Empty, CCGetRequestParam("SiteAuthorizeDescription", Method))
        Set SiteProPrice = CCCreateControl(ccsTextBox, "SiteProPrice", "Pro Price", ccsFloat, Empty, CCGetRequestParam("SiteProPrice", Method))
        Set SiteUseEditor = CCCreateControl(ccsCheckBox, "SiteUseEditor", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("SiteUseEditor", Method))
        Set SiteDebugMode = CCCreateControl(ccsCheckBox, "SiteDebugMode", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("SiteDebugMode", Method))
        Set SiteDebugIP = CCCreateControl(ccsTextArea, "SiteDebugIP", "Debug IP", ccsMemo, Empty, CCGetRequestParam("SiteDebugIP", Method))
        Set SiteDebugModeCookieTTL = CCCreateControl(ccsTextBox, "SiteDebugModeCookieTTL", "Debug Mode Cookie TTL", ccsInteger, Empty, CCGetRequestParam("SiteDebugModeCookieTTL", Method))
        Set DebugDomainName = CCCreateControl(ccsLabel, "DebugDomainName", Empty, ccsText, Empty, CCGetRequestParam("DebugDomainName", Method))
        Set SiteContentManagementIP = CCCreateControl(ccsTextArea, "SiteContentManagementIP", "Content Management IP", ccsMemo, Empty, CCGetRequestParam("SiteContentManagementIP", Method))
        Set SiteContentManagementModeCookieTTL = CCCreateControl(ccsTextBox, "SiteContentManagementModeCookieTTL", "Content Management Cookie TTL", ccsInteger, Empty, CCGetRequestParam("SiteContentManagementModeCookieTTL", Method))
        Set SiteRedirectPageOnError = CCCreateControl(ccsTextBox, "SiteRedirectPageOnError", "Redirect Page On Error", ccsText, Empty, CCGetRequestParam("SiteRedirectPageOnError", Method))
        Set SiteLoginPage = CCCreateControl(ccsTextBox, "SiteLoginPage", "Login Page", ccsText, Empty, CCGetRequestParam("SiteLoginPage", Method))
        Set SiteUploadPath = CCCreateControl(ccsTextBox, "SiteUploadPath", "Upload Path", ccsText, Empty, CCGetRequestParam("SiteUploadPath", Method))
        Set SitePageTemplateExecuteASPPath = CCCreateControl(ccsTextBox, "SitePageTemplateExecuteASPPath", "Page Template Execute ASPPath", ccsText, Empty, CCGetRequestParam("SitePageTemplateExecuteASPPath", Method))
        Set SiteCustom404ScriptName = CCCreateControl(ccsTextBox, "SiteCustom404ScriptName", "Custom 404 Script Name", ccsText, Empty, CCGetRequestParam("SiteCustom404ScriptName", Method))
        Set Button_Insert = CCCreateButton("Button_Insert", Method)
        Set Button_Update = CCCreateButton("Button_Update", Method)
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(SiteName, SiteDomainName, SiteDescription, SiteMaxTemplateBlocks, SiteMainTemplateID, SiteURLEncryptionKey, SiteDataEncryptionKey, _
             SiteTrackingString, SiteTrackingLinkLength, SiteTrackingCookieTTL, SiteDefaultURL, SiteRefererCookieTTL, SiteDefaultPageTemplateID, SiteAuthorization, SiteAuthorizeLogin, _
             SiteAuthorizeTransactionKey, SiteAuthorizeDescription, SiteProPrice, SiteUseEditor, SiteDebugMode, SiteDebugIP, SiteDebugModeCookieTTL, SiteContentManagementIP, _
             SiteContentManagementModeCookieTTL, SiteRedirectPageOnError, SiteLoginPage, SiteUploadPath, SitePageTemplateExecuteASPPath, SiteCustom404ScriptName)
    End Sub
'End Site Class_Initialize Event

'Site Initialize Method @55-A5E37F03
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("sesSiteID") = Session("SiteID")
        End With
    End Sub
'End Site Initialize Method

'Site Class_Terminate Event @55-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End Site Class_Terminate Event

'Site Validate Method @55-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End Site Validate Method

'Site Operation Method @55-EEBBED02
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = IIf(EditMode, "Button_Update", "Button_Insert")
            If Button_Insert.Pressed Then
                PressedButton = "Button_Insert"
            ElseIf Button_Update.Pressed Then
                PressedButton = "Button_Update"
            ElseIf Button_Cancel.Pressed Then
                PressedButton = "Button_Cancel"
            End If
        End If
        Redirect = "AdminSiteSettings.asp"
        If PressedButton = "Button_Cancel" Then
            If NOT Button_Cancel.OnClick Then
                Redirect = ""
            End If
        ElseIf Validate() Then
            If PressedButton = "Button_Insert" Then
                If NOT Button_Insert.OnClick() OR NOT InsertRow() Then
                    Redirect = ""
                End If
            ElseIf PressedButton = "Button_Update" Then
                If NOT Button_Update.OnClick() OR NOT UpdateRow() Then
                    Redirect = ""
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End Site Operation Method

'Site InsertRow Method @55-8DF28435
    Function InsertRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeInsert", Me)
        If NOT InsertAllowed Then InsertRow = False : Exit Function
        DataSource.SiteName.Value = SiteName.Value
        DataSource.SiteDomainName.Value = SiteDomainName.Value
        DataSource.SiteDescription.Value = SiteDescription.Value
        DataSource.SiteMaxTemplateBlocks.Value = SiteMaxTemplateBlocks.Value
        DataSource.SiteMainTemplateID.Value = SiteMainTemplateID.Value
        DataSource.SiteURLEncryptionKey.Value = SiteURLEncryptionKey.Value
        DataSource.SiteDataEncryptionKey.Value = SiteDataEncryptionKey.Value
        DataSource.SiteTrackingString.Value = SiteTrackingString.Value
        DataSource.SiteTrackingLinkLength.Value = SiteTrackingLinkLength.Value
        DataSource.SiteTrackingCookieTTL.Value = SiteTrackingCookieTTL.Value
        DataSource.SiteDefaultURL.Value = SiteDefaultURL.Value
        DataSource.SiteRefererCookieTTL.Value = SiteRefererCookieTTL.Value
        DataSource.SiteDefaultPageTemplateID.Value = SiteDefaultPageTemplateID.Value
        DataSource.SiteAuthorization.Value = SiteAuthorization.Value
        DataSource.SiteAuthorizeLogin.Value = SiteAuthorizeLogin.Value
        DataSource.SiteAuthorizeTransactionKey.Value = SiteAuthorizeTransactionKey.Value
        DataSource.SiteAuthorizeDescription.Value = SiteAuthorizeDescription.Value
        DataSource.SiteProPrice.Value = SiteProPrice.Value
        DataSource.SiteUseEditor.Value = SiteUseEditor.Value
        DataSource.SiteDebugMode.Value = SiteDebugMode.Value
        DataSource.SiteDebugIP.Value = SiteDebugIP.Value
        DataSource.SiteDebugModeCookieTTL.Value = SiteDebugModeCookieTTL.Value
        DataSource.SiteContentManagementIP.Value = SiteContentManagementIP.Value
        DataSource.SiteContentManagementModeCookieTTL.Value = SiteContentManagementModeCookieTTL.Value
        DataSource.SiteRedirectPageOnError.Value = SiteRedirectPageOnError.Value
        DataSource.SiteLoginPage.Value = SiteLoginPage.Value
        DataSource.SiteUploadPath.Value = SiteUploadPath.Value
        DataSource.SitePageTemplateExecuteASPPath.Value = SitePageTemplateExecuteASPPath.Value
        DataSource.SiteCustom404ScriptName.Value = SiteCustom404ScriptName.Value
        DataSource.Insert(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInsert", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        InsertRow = (Errors.Count = 0)
    End Function
'End Site InsertRow Method

'Site UpdateRow Method @55-2C26678E
    Function UpdateRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUpdate", Me)
        If NOT UpdateAllowed Then UpdateRow = False : Exit Function
        DataSource.SiteName.Value = SiteName.Value
        DataSource.SiteDomainName.Value = SiteDomainName.Value
        DataSource.SiteDescription.Value = SiteDescription.Value
        DataSource.SiteMaxTemplateBlocks.Value = SiteMaxTemplateBlocks.Value
        DataSource.SiteMainTemplateID.Value = SiteMainTemplateID.Value
        DataSource.SiteURLEncryptionKey.Value = SiteURLEncryptionKey.Value
        DataSource.SiteDataEncryptionKey.Value = SiteDataEncryptionKey.Value
        DataSource.SiteTrackingString.Value = SiteTrackingString.Value
        DataSource.SiteTrackingLinkLength.Value = SiteTrackingLinkLength.Value
        DataSource.SiteTrackingCookieTTL.Value = SiteTrackingCookieTTL.Value
        DataSource.SiteDefaultURL.Value = SiteDefaultURL.Value
        DataSource.SiteRefererCookieTTL.Value = SiteRefererCookieTTL.Value
        DataSource.SiteDefaultPageTemplateID.Value = SiteDefaultPageTemplateID.Value
        DataSource.SiteAuthorization.Value = SiteAuthorization.Value
        DataSource.SiteAuthorizeLogin.Value = SiteAuthorizeLogin.Value
        DataSource.SiteAuthorizeTransactionKey.Value = SiteAuthorizeTransactionKey.Value
        DataSource.SiteAuthorizeDescription.Value = SiteAuthorizeDescription.Value
        DataSource.SiteProPrice.Value = SiteProPrice.Value
        DataSource.SiteUseEditor.Value = SiteUseEditor.Value
        DataSource.SiteDebugMode.Value = SiteDebugMode.Value
        DataSource.SiteDebugIP.Value = SiteDebugIP.Value
        DataSource.SiteDebugModeCookieTTL.Value = SiteDebugModeCookieTTL.Value
        DataSource.SiteContentManagementIP.Value = SiteContentManagementIP.Value
        DataSource.SiteContentManagementModeCookieTTL.Value = SiteContentManagementModeCookieTTL.Value
        DataSource.SiteRedirectPageOnError.Value = SiteRedirectPageOnError.Value
        DataSource.SiteLoginPage.Value = SiteLoginPage.Value
        DataSource.SiteUploadPath.Value = SiteUploadPath.Value
        DataSource.SitePageTemplateExecuteASPPath.Value = SitePageTemplateExecuteASPPath.Value
        DataSource.SiteCustom404ScriptName.Value = SiteCustom404ScriptName.Value
        DataSource.Update(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterUpdate", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        UpdateRow = (Errors.Count = 0)
    End Function
'End Site UpdateRow Method

'Site Show Method @55-5DD02B6F
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "Site" & IIf(EditMode, ":Edit", ""))
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
            Array(SiteName,  SiteDomainName,  SiteDescription,  SiteMaxTemplateBlocks,  SiteMainTemplateID,  SiteURLEncryptionKey,  SiteDataEncryptionKey, _
                 SiteTrackingString,  SiteTrackingLinkLength,  SiteTrackingCookieTTL,  SiteDefaultURL,  SiteRefererCookieTTL,  SiteDefaultPageTemplateID,  SiteAuthorization,  SiteAuthorizeLogin, _
                 SiteAuthorizeTransactionKey,  SiteAuthorizeDescription,  SiteProPrice,  SiteUseEditor,  SiteDebugMode,  SiteDebugIP,  SiteDebugModeCookieTTL,  DebugDomainName, _
                 SiteContentManagementIP,  SiteContentManagementModeCookieTTL,  SiteRedirectPageOnError,  SiteLoginPage,  SiteUploadPath,  SitePageTemplateExecuteASPPath,  SiteCustom404ScriptName,  Button_Insert,  Button_Update,  Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        SiteName.Value = Recordset.Fields("SiteName")
                        SiteDomainName.Value = Recordset.Fields("SiteDomainName")
                        SiteDescription.Value = Recordset.Fields("SiteDescription")
                        SiteMaxTemplateBlocks.Value = Recordset.Fields("SiteMaxTemplateBlocks")
                        SiteMainTemplateID.Value = Recordset.Fields("SiteMainTemplateID")
                        SiteURLEncryptionKey.Value = Recordset.Fields("SiteURLEncryptionKey")
                        SiteDataEncryptionKey.Value = Recordset.Fields("SiteDataEncryptionKey")
                        SiteTrackingString.Value = Recordset.Fields("SiteTrackingString")
                        SiteTrackingLinkLength.Value = Recordset.Fields("SiteTrackingLinkLength")
                        SiteTrackingCookieTTL.Value = Recordset.Fields("SiteTrackingCookieTTL")
                        SiteDefaultURL.Value = Recordset.Fields("SiteDefaultURL")
                        SiteRefererCookieTTL.Value = Recordset.Fields("SiteRefererCookieTTL")
                        SiteDefaultPageTemplateID.Value = Recordset.Fields("SiteDefaultPageTemplateID")
                        SiteAuthorization.Value = Recordset.Fields("SiteAuthorization")
                        SiteAuthorizeLogin.Value = Recordset.Fields("SiteAuthorizeLogin")
                        SiteAuthorizeTransactionKey.Value = Recordset.Fields("SiteAuthorizeTransactionKey")
                        SiteAuthorizeDescription.Value = Recordset.Fields("SiteAuthorizeDescription")
                        SiteProPrice.Value = Recordset.Fields("SiteProPrice")
                        SiteUseEditor.Value = Recordset.Fields("SiteUseEditor")
                        SiteDebugMode.Value = Recordset.Fields("SiteDebugMode")
                        SiteDebugIP.Value = Recordset.Fields("SiteDebugIP")
                        SiteDebugModeCookieTTL.Value = Recordset.Fields("SiteDebugModeCookieTTL")
                        SiteContentManagementIP.Value = Recordset.Fields("SiteContentManagementIP")
                        SiteContentManagementModeCookieTTL.Value = Recordset.Fields("SiteContentManagementModeCookieTTL")
                        SiteRedirectPageOnError.Value = Recordset.Fields("SiteRedirectPageOnError")
                        SiteLoginPage.Value = Recordset.Fields("SiteLoginPage")
                        SiteUploadPath.Value = Recordset.Fields("SiteUploadPath")
                        SitePageTemplateExecuteASPPath.Value = Recordset.Fields("SitePageTemplateExecuteASPPath")
                        SiteCustom404ScriptName.Value = Recordset.Fields("SiteCustom404ScriptName")
                    End If
                Else
                    EditMode = False
                End If
            End If
            If EditMode Then
                
            End If
        End If
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors SiteName.Errors
            Errors.AddErrors SiteDomainName.Errors
            Errors.AddErrors SiteDescription.Errors
            Errors.AddErrors SiteMaxTemplateBlocks.Errors
            Errors.AddErrors SiteMainTemplateID.Errors
            Errors.AddErrors SiteURLEncryptionKey.Errors
            Errors.AddErrors SiteDataEncryptionKey.Errors
            Errors.AddErrors SiteTrackingString.Errors
            Errors.AddErrors SiteTrackingLinkLength.Errors
            Errors.AddErrors SiteTrackingCookieTTL.Errors
            Errors.AddErrors SiteDefaultURL.Errors
            Errors.AddErrors SiteRefererCookieTTL.Errors
            Errors.AddErrors SiteDefaultPageTemplateID.Errors
            Errors.AddErrors SiteAuthorization.Errors
            Errors.AddErrors SiteAuthorizeLogin.Errors
            Errors.AddErrors SiteAuthorizeTransactionKey.Errors
            Errors.AddErrors SiteAuthorizeDescription.Errors
            Errors.AddErrors SiteProPrice.Errors
            Errors.AddErrors SiteUseEditor.Errors
            Errors.AddErrors SiteDebugMode.Errors
            Errors.AddErrors SiteDebugIP.Errors
            Errors.AddErrors SiteDebugModeCookieTTL.Errors
            Errors.AddErrors SiteContentManagementIP.Errors
            Errors.AddErrors SiteContentManagementModeCookieTTL.Errors
            Errors.AddErrors SiteRedirectPageOnError.Errors
            Errors.AddErrors SiteLoginPage.Errors
            Errors.AddErrors SiteUploadPath.Errors
            Errors.AddErrors SitePageTemplateExecuteASPPath.Errors
            Errors.AddErrors SiteCustom404ScriptName.Errors
            Errors.AddErrors DataSource.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction
        Button_Insert.Visible = NOT EditMode AND InsertAllowed
        Button_Update.Visible = EditMode AND UpdateAllowed

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End Site Show Method

End Class 'End Site Class @55-A61BA892

Class clsSiteDataSource 'SiteDataSource Class @55-F0A0929C

'DataSource Variables @55-6880F735
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
    Public SiteName
    Public SiteDomainName
    Public SiteDescription
    Public SiteMaxTemplateBlocks
    Public SiteMainTemplateID
    Public SiteURLEncryptionKey
    Public SiteDataEncryptionKey
    Public SiteTrackingString
    Public SiteTrackingLinkLength
    Public SiteTrackingCookieTTL
    Public SiteDefaultURL
    Public SiteRefererCookieTTL
    Public SiteDefaultPageTemplateID
    Public SiteAuthorization
    Public SiteAuthorizeLogin
    Public SiteAuthorizeTransactionKey
    Public SiteAuthorizeDescription
    Public SiteProPrice
    Public SiteUseEditor
    Public SiteDebugMode
    Public SiteDebugIP
    Public SiteDebugModeCookieTTL
    Public SiteContentManagementIP
    Public SiteContentManagementModeCookieTTL
    Public SiteRedirectPageOnError
    Public SiteLoginPage
    Public SiteUploadPath
    Public SitePageTemplateExecuteASPPath
    Public SiteCustom404ScriptName
'End DataSource Variables

'DataSource Class_Initialize Event @55-E584EFCE
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set SiteName = CCCreateField("SiteName", "SiteName", ccsText, Empty, Recordset)
        Set SiteDomainName = CCCreateField("SiteDomainName", "SiteDomainName", ccsText, Empty, Recordset)
        Set SiteDescription = CCCreateField("SiteDescription", "SiteDescription", ccsMemo, Empty, Recordset)
        Set SiteMaxTemplateBlocks = CCCreateField("SiteMaxTemplateBlocks", "SiteMaxTemplateBlocks", ccsInteger, Empty, Recordset)
        Set SiteMainTemplateID = CCCreateField("SiteMainTemplateID", "SiteMainTemplateID", ccsText, Empty, Recordset)
        Set SiteURLEncryptionKey = CCCreateField("SiteURLEncryptionKey", "SiteURLEncryptionKey", ccsText, Empty, Recordset)
        Set SiteDataEncryptionKey = CCCreateField("SiteDataEncryptionKey", "SiteDataEncryptionKey", ccsText, Empty, Recordset)
        Set SiteTrackingString = CCCreateField("SiteTrackingString", "SiteTrackingString", ccsText, Empty, Recordset)
        Set SiteTrackingLinkLength = CCCreateField("SiteTrackingLinkLength", "SiteTrackingLinkLength", ccsInteger, Empty, Recordset)
        Set SiteTrackingCookieTTL = CCCreateField("SiteTrackingCookieTTL", "SiteTrackingCookieTTL", ccsInteger, Empty, Recordset)
        Set SiteDefaultURL = CCCreateField("SiteDefaultURL", "SiteDefaultURL", ccsText, Empty, Recordset)
        Set SiteRefererCookieTTL = CCCreateField("SiteRefererCookieTTL", "SiteRefererCookieTTL", ccsInteger, Empty, Recordset)
        Set SiteDefaultPageTemplateID = CCCreateField("SiteDefaultPageTemplateID", "SiteDefaultPageTemplateID", ccsText, Empty, Recordset)
        Set SiteAuthorization = CCCreateField("SiteAuthorization", "SiteAuthorization", ccsText, Empty, Recordset)
        Set SiteAuthorizeLogin = CCCreateField("SiteAuthorizeLogin", "SiteAuthorizeLogin", ccsText, Empty, Recordset)
        Set SiteAuthorizeTransactionKey = CCCreateField("SiteAuthorizeTransactionKey", "SiteAuthorizeTransactionKey", ccsText, Empty, Recordset)
        Set SiteAuthorizeDescription = CCCreateField("SiteAuthorizeDescription", "SiteAuthorizeDescription", ccsText, Empty, Recordset)
        Set SiteProPrice = CCCreateField("SiteProPrice", "SiteProPrice", ccsFloat, Empty, Recordset)
        Set SiteUseEditor = CCCreateField("SiteUseEditor", "SiteUseEditor", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set SiteDebugMode = CCCreateField("SiteDebugMode", "SiteDebugMode", ccsBoolean, Array(1, 0, Empty), Recordset)
        Set SiteDebugIP = CCCreateField("SiteDebugIP", "SiteDebugIP", ccsMemo, Empty, Recordset)
        Set SiteDebugModeCookieTTL = CCCreateField("SiteDebugModeCookieTTL", "SiteDebugModeCookieTTL", ccsInteger, Empty, Recordset)
        Set SiteContentManagementIP = CCCreateField("SiteContentManagementIP", "SiteContentManagementIP", ccsMemo, Empty, Recordset)
        Set SiteContentManagementModeCookieTTL = CCCreateField("SiteContentManagementModeCookieTTL", "SiteContentManagementModeCookieTTL", ccsInteger, Empty, Recordset)
        Set SiteRedirectPageOnError = CCCreateField("SiteRedirectPageOnError", "SiteRedirectPageOnError", ccsText, Empty, Recordset)
        Set SiteLoginPage = CCCreateField("SiteLoginPage", "SiteLoginPage", ccsText, Empty, Recordset)
        Set SiteUploadPath = CCCreateField("SiteUploadPath", "SiteUploadPath", ccsText, Empty, Recordset)
        Set SitePageTemplateExecuteASPPath = CCCreateField("SitePageTemplateExecuteASPPath", "SitePageTemplateExecuteASPPath", ccsText, Empty, Recordset)
        Set SiteCustom404ScriptName = CCCreateField("SiteCustom404ScriptName", "SiteCustom404ScriptName", ccsText, Empty, Recordset)
        Fields.AddFields Array(SiteName,  SiteDomainName,  SiteDescription,  SiteMaxTemplateBlocks,  SiteMainTemplateID,  SiteURLEncryptionKey,  SiteDataEncryptionKey, _
             SiteTrackingString,  SiteTrackingLinkLength,  SiteTrackingCookieTTL,  SiteDefaultURL,  SiteRefererCookieTTL,  SiteDefaultPageTemplateID,  SiteAuthorization,  SiteAuthorizeLogin, _
             SiteAuthorizeTransactionKey,  SiteAuthorizeDescription,  SiteProPrice,  SiteUseEditor,  SiteDebugMode,  SiteDebugIP,  SiteDebugModeCookieTTL,  SiteContentManagementIP, _
             SiteContentManagementModeCookieTTL,  SiteRedirectPageOnError,  SiteLoginPage,  SiteUploadPath,  SitePageTemplateExecuteASPPath,  SiteCustom404ScriptName)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1  *  " & vbLf & _
        "FROM Site {SQL_Where} {SQL_OrderBy}"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @55-D78877EB
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
            AllParamsSet = .AllParamsSet
            .Criterion(1) = .Operation(opEqual, False, "[SiteID]", .getParamByID(1))
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

'Open Method @55-48A2DA7D
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
        BuildTableWhere
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

'DataSource Class_Terminate Event @55-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

'Update Method @55-1D1C6424
    Sub Update(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildUpdate", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Cmd.Prepared = True
        BuildTableWhere
        If NOT AllParamsSet Then
            Errors.AddError(CCSLocales.GetText("CCS_CustomOperationError_MissingParameters", Empty))
        End If
        Cmd.SQL = "UPDATE [Site] SET " & _
            "[SiteName]=" & Connection.ToSQL(SiteName, SiteName.DataType) & ", " & _
            "[SiteDomainName]=" & Connection.ToSQL(SiteDomainName, SiteDomainName.DataType) & ", " & _
            "[SiteDescription]=?" & ", " & _
            "[SiteMaxTemplateBlocks]=" & Connection.ToSQL(SiteMaxTemplateBlocks, SiteMaxTemplateBlocks.DataType) & ", " & _
            "[SiteMainTemplateID]=" & Connection.ToSQL(SiteMainTemplateID, SiteMainTemplateID.DataType) & ", " & _
            "[SiteURLEncryptionKey]=" & Connection.ToSQL(SiteURLEncryptionKey, SiteURLEncryptionKey.DataType) & ", " & _
            "[SiteDataEncryptionKey]=" & Connection.ToSQL(SiteDataEncryptionKey, SiteDataEncryptionKey.DataType) & ", " & _
            "[SiteTrackingString]=" & Connection.ToSQL(SiteTrackingString, SiteTrackingString.DataType) & ", " & _
            "[SiteTrackingLinkLength]=" & Connection.ToSQL(SiteTrackingLinkLength, SiteTrackingLinkLength.DataType) & ", " & _
            "[SiteTrackingCookieTTL]=" & Connection.ToSQL(SiteTrackingCookieTTL, SiteTrackingCookieTTL.DataType) & ", " & _
            "[SiteDefaultURL]=" & Connection.ToSQL(SiteDefaultURL, SiteDefaultURL.DataType) & ", " & _
            "[SiteRefererCookieTTL]=" & Connection.ToSQL(SiteRefererCookieTTL, SiteRefererCookieTTL.DataType) & ", " & _
            "[SiteDefaultPageTemplateID]=" & Connection.ToSQL(SiteDefaultPageTemplateID, SiteDefaultPageTemplateID.DataType) & ", " & _
            "[SiteAuthorization]=" & Connection.ToSQL(SiteAuthorization, SiteAuthorization.DataType) & ", " & _
            "[SiteAuthorizeLogin]=" & Connection.ToSQL(SiteAuthorizeLogin, SiteAuthorizeLogin.DataType) & ", " & _
            "[SiteAuthorizeTransactionKey]=" & Connection.ToSQL(SiteAuthorizeTransactionKey, SiteAuthorizeTransactionKey.DataType) & ", " & _
            "[SiteAuthorizeDescription]=" & Connection.ToSQL(SiteAuthorizeDescription, SiteAuthorizeDescription.DataType) & ", " & _
            "[SiteProPrice]=" & Connection.ToSQL(SiteProPrice, SiteProPrice.DataType) & ", " & _
            "[SiteUseEditor]=" & Connection.ToSQL(SiteUseEditor, SiteUseEditor.DataType) & ", " & _
            "[SiteDebugMode]=" & Connection.ToSQL(SiteDebugMode, SiteDebugMode.DataType) & ", " & _
            "[SiteDebugIP]=?" & ", " & _
            "[SiteDebugModeCookieTTL]=" & Connection.ToSQL(SiteDebugModeCookieTTL, SiteDebugModeCookieTTL.DataType) & ", " & _
            "[SiteContentManagementIP]=?" & ", " & _
            "[SiteContentManagementModeCookieTTL]=" & Connection.ToSQL(SiteContentManagementModeCookieTTL, SiteContentManagementModeCookieTTL.DataType) & ", " & _
            "[SiteRedirectPageOnError]=" & Connection.ToSQL(SiteRedirectPageOnError, SiteRedirectPageOnError.DataType) & ", " & _
            "[SiteLoginPage]=" & Connection.ToSQL(SiteLoginPage, SiteLoginPage.DataType) & ", " & _
            "[SiteUploadPath]=" & Connection.ToSQL(SiteUploadPath, SiteUploadPath.DataType) & ", " & _
            "[SitePageTemplateExecuteASPPath]=" & Connection.ToSQL(SitePageTemplateExecuteASPPath, SitePageTemplateExecuteASPPath.DataType) & ", " & _
            "[SiteCustom404ScriptName]=" & Connection.ToSQL(SiteCustom404ScriptName, SiteCustom404ScriptName.DataType) & _
            IIf(Len(Where) > 0, " WHERE " & Where, "")
        Cmd.CommandParameters = Array( _
            Array("[SiteDescription]", adLongVarChar, adParamInput, 2147483647, SiteDescription.Value), _
            Array("[SiteDebugIP]", adLongVarChar, adParamInput, 2147483647, SiteDebugIP.Value), _
            Array("[SiteContentManagementIP]", adLongVarChar, adParamInput, 2147483647, SiteContentManagementIP.Value))
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

'Insert Method @55-F2C71F9B
    Sub Insert(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildInsert", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Cmd.Prepared = True
        Cmd.SQL = "INSERT INTO [Site] (" & _
            "[SiteName], " & _
            "[SiteDomainName], " & _
            "[SiteDescription], " & _
            "[SiteMaxTemplateBlocks], " & _
            "[SiteMainTemplateID], " & _
            "[SiteURLEncryptionKey], " & _
            "[SiteDataEncryptionKey], " & _
            "[SiteTrackingString], " & _
            "[SiteTrackingLinkLength], " & _
            "[SiteTrackingCookieTTL], " & _
            "[SiteDefaultURL], " & _
            "[SiteRefererCookieTTL], " & _
            "[SiteDefaultPageTemplateID], " & _
            "[SiteAuthorization], " & _
            "[SiteAuthorizeLogin], " & _
            "[SiteAuthorizeTransactionKey], " & _
            "[SiteAuthorizeDescription], " & _
            "[SiteProPrice], " & _
            "[SiteUseEditor], " & _
            "[SiteDebugMode], " & _
            "[SiteDebugIP], " & _
            "[SiteDebugModeCookieTTL], " & _
            "[SiteContentManagementIP], " & _
            "[SiteContentManagementModeCookieTTL], " & _
            "[SiteRedirectPageOnError], " & _
            "[SiteLoginPage], " & _
            "[SiteUploadPath], " & _
            "[SitePageTemplateExecuteASPPath], " & _
            "[SiteCustom404ScriptName]" & _
        ") VALUES (" & _
            Connection.ToSQL(SiteName, SiteName.DataType) & ", " & _
            Connection.ToSQL(SiteDomainName, SiteDomainName.DataType) & ", " & _
            "?" & ", " & _
            Connection.ToSQL(SiteMaxTemplateBlocks, SiteMaxTemplateBlocks.DataType) & ", " & _
            Connection.ToSQL(SiteMainTemplateID, SiteMainTemplateID.DataType) & ", " & _
            Connection.ToSQL(SiteURLEncryptionKey, SiteURLEncryptionKey.DataType) & ", " & _
            Connection.ToSQL(SiteDataEncryptionKey, SiteDataEncryptionKey.DataType) & ", " & _
            Connection.ToSQL(SiteTrackingString, SiteTrackingString.DataType) & ", " & _
            Connection.ToSQL(SiteTrackingLinkLength, SiteTrackingLinkLength.DataType) & ", " & _
            Connection.ToSQL(SiteTrackingCookieTTL, SiteTrackingCookieTTL.DataType) & ", " & _
            Connection.ToSQL(SiteDefaultURL, SiteDefaultURL.DataType) & ", " & _
            Connection.ToSQL(SiteRefererCookieTTL, SiteRefererCookieTTL.DataType) & ", " & _
            Connection.ToSQL(SiteDefaultPageTemplateID, SiteDefaultPageTemplateID.DataType) & ", " & _
            Connection.ToSQL(SiteAuthorization, SiteAuthorization.DataType) & ", " & _
            Connection.ToSQL(SiteAuthorizeLogin, SiteAuthorizeLogin.DataType) & ", " & _
            Connection.ToSQL(SiteAuthorizeTransactionKey, SiteAuthorizeTransactionKey.DataType) & ", " & _
            Connection.ToSQL(SiteAuthorizeDescription, SiteAuthorizeDescription.DataType) & ", " & _
            Connection.ToSQL(SiteProPrice, SiteProPrice.DataType) & ", " & _
            Connection.ToSQL(SiteUseEditor, SiteUseEditor.DataType) & ", " & _
            Connection.ToSQL(SiteDebugMode, SiteDebugMode.DataType) & ", " & _
            "?" & ", " & _
            Connection.ToSQL(SiteDebugModeCookieTTL, SiteDebugModeCookieTTL.DataType) & ", " & _
            "?" & ", " & _
            Connection.ToSQL(SiteContentManagementModeCookieTTL, SiteContentManagementModeCookieTTL.DataType) & ", " & _
            Connection.ToSQL(SiteRedirectPageOnError, SiteRedirectPageOnError.DataType) & ", " & _
            Connection.ToSQL(SiteLoginPage, SiteLoginPage.DataType) & ", " & _
            Connection.ToSQL(SiteUploadPath, SiteUploadPath.DataType) & ", " & _
            Connection.ToSQL(SitePageTemplateExecuteASPPath, SitePageTemplateExecuteASPPath.DataType) & ", " & _
            Connection.ToSQL(SiteCustom404ScriptName, SiteCustom404ScriptName.DataType) & _
        ")"
        Cmd.CommandParameters = Array( _
            Array("[SiteDescription]", adLongVarChar, adParamInput,2147483647, SiteDescription.Value),  _
            Array("[SiteDebugIP]", adLongVarChar, adParamInput,2147483647, SiteDebugIP.Value),  _
            Array("[SiteContentManagementIP]", adLongVarChar, adParamInput,2147483647, SiteContentManagementIP.Value) _
        )
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteInsert", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteInsert", Me)
        End If
    End Sub
'End Insert Method

End Class 'End SiteDataSource Class @55-A61BA892


%>
