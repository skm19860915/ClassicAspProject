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

'Initialize Page @1-2EC1CD09
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
Dim Menu
Dim SiteID
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminDashboard.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminDashboard.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-37EDAD5A

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set SiteID = CCCreateControl(ccsHidden, "SiteID", Empty, ccsInteger, Empty, CCGetRequestParam("SiteID", ccsGet))
Menu.Value = DHTMLMenu
If IsEmpty(SiteID.Value) Then _
    SiteID.Value = Session("SiteID")



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

'Show Page @1-D1953559
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, SiteID))
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

'UnloadPage Sub @1-E05A2E62
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    Set CCSEvents = Nothing
End Sub
'End UnloadPage Sub


%>
