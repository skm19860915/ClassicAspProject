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

'Initialize Page @1-E1F83926
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

' Page controls
Dim Login
Dim ChildControls


Redirect = ""
'Redirect = "AdminDashboard.asp"
TemplateFileName = "Login.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "Login.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page
'Initialize Objects @1-0C5F0BDF

' Controls
Set Login = new clsRecordLogin

' Events
%>
<!-- #INCLUDE FILE="Login_events.asp" -->
<%
BindEvents


CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-6AA0DE03
Login.Operation
'End Execute Components



'response.write(redirect)
'response.End
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

'Show Page @1-87B9E5D5
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Login))
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

'UnloadPage Sub @1-7E60FA98
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    Set CCSEvents = Nothing
    Set Login = Nothing
End Sub
'End UnloadPage Sub



Class clsRecordLogin 'Login Class @2-D3C99874

'Login Variables @2-CC240756

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
    Dim login
    Dim password
    Dim Button_DoLogin
'End Login Variables

'Login Class_Initialize Event @2-023B96FC
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
            FormSubmitted = (OperationMode(0) = "Login")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "Login"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set login = CCCreateControl(ccsTextBox, "login", Empty, ccsText, Empty, CCGetRequestParam("login", Method))
        login.Required = True
        Set password = CCCreateControl(ccsTextBox, "password", Empty, ccsText, Empty, CCGetRequestParam("password", Method))
        password.Required = True
        Set Button_DoLogin = CCCreateButton("Button_DoLogin", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(login, password)
    End Sub
'End Login Class_Initialize Event

'Login Class_Terminate Event @2-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End Login Class_Terminate Event

'Login Validate Method @2-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End Login Validate Method

'Login Operation Method @2-8700D4AE
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = "Button_DoLogin"
            If Button_DoLogin.Pressed Then
                PressedButton = "Button_DoLogin"
            End If
        End If
        Redirect = FileName & ""
        If Validate() Then
            If PressedButton = "Button_DoLogin" Then
                If NOT Button_DoLogin.OnClick() Then
                    Redirect = ""
                End If
            End If
        Else
            Redirect = ""
        End If
    End Sub
'End Login Operation Method

'Login Show Method @2-4CF16042
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        EditMode = False
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "Login" & IIf(EditMode, ":Edit", ""))
        Set TemplateBlock = Tpl.Block("Record " & ComponentName)
        If TemplateBlock is Nothing Then Exit Sub
        TemplateBlock.Variable("HTMLFormName") = ComponentName
        TemplateBlock.Variable("HTMLFormEnctype") ="application/x-www-form-urlencoded"
        Set Controls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(login, password, Button_DoLogin))
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors login.Errors
            Errors.AddErrors password.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End Login Show Method

End Class 'End Login Class @2-A61BA892


%>
