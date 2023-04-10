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

'Initialize Page @1-005A9A06
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
Dim PageTemplateSection
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminPageTemplateSectionEdit.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminPageTemplateSectionEdit.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Initialize Objects @1-A2888B75
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set PageTemplateSection = new clsRecordPageTemplateSection
Menu.Value = DHTMLMenu

PageTemplateSection.Initialize DBSystem

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-3F5C0ECA
PageTemplateSection.Operation
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

'Show Page @1-A6EF56B3
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, PageTemplateSection))
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

'UnloadPage Sub @1-6E4270BA
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set PageTemplateSection = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordPageTemplateSection 'PageTemplateSection Class @33-9A6C5EC3

'PageTemplateSection Variables @33-D257BB3D

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
    Dim PageTemplateSectionNickname
    Dim PageTemplateSectionSiteID
    Dim PageTemplateSectionName
    Dim PageTemplateSectionDescription
    Dim Button_Insert
    Dim Button_Update
    Dim Button_Delete
    Dim Button_Cancel
'End PageTemplateSection Variables

'PageTemplateSection Class_Initialize Event @33-D2695BA8
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsPageTemplateSectionDataSource
        Set Command = New clsCommand
        InsertAllowed = True
        UpdateAllowed = True
        DeleteAllowed = True
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "PageTemplateSection")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "PageTemplateSection"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set PageTemplateSectionNickname = CCCreateControl(ccsTextBox, "PageTemplateSectionNickname", "Section Nickname", ccsText, Empty, CCGetRequestParam("PageTemplateSectionNickname", Method))
        PageTemplateSectionNickname.Required = True
        Set PageTemplateSectionSiteID = CCCreateControl(ccsHidden, "PageTemplateSectionSiteID", Empty, ccsInteger, Empty, CCGetRequestParam("PageTemplateSectionSiteID", Method))
        Set PageTemplateSectionName = CCCreateControl(ccsTextBox, "PageTemplateSectionName", "Name", ccsText, Empty, CCGetRequestParam("PageTemplateSectionName", Method))
        Set PageTemplateSectionDescription = CCCreateControl(ccsTextArea, "PageTemplateSectionDescription", "Description", ccsMemo, Empty, CCGetRequestParam("PageTemplateSectionDescription", Method))
        Set Button_Insert = CCCreateButton("Button_Insert", Method)
        Set Button_Update = CCCreateButton("Button_Update", Method)
        Set Button_Delete = CCCreateButton("Button_Delete", Method)
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(PageTemplateSectionNickname, PageTemplateSectionSiteID, PageTemplateSectionName, PageTemplateSectionDescription)
        If Not FormSubmitted Then
            If IsEmpty(PageTemplateSectionSiteID.Value) Then _
                PageTemplateSectionSiteID.Value = Session("SiteID")
        End If
    End Sub
'End PageTemplateSection Class_Initialize Event

'PageTemplateSection Initialize Method @33-29560EF5
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("urlPageTemplateSectionID") = CCGetRequestParam("PageTemplateSectionID", ccsGET)
        End With
    End Sub
'End PageTemplateSection Initialize Method

'PageTemplateSection Class_Terminate Event @33-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End PageTemplateSection Class_Terminate Event

'PageTemplateSection Validate Method @33-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End PageTemplateSection Validate Method

'PageTemplateSection Operation Method @33-8283DF94
    Sub Operation()
        If NOT ( Visible AND FormSubmitted ) Then Exit Sub

        If FormSubmitted Then
            PressedButton = IIf(EditMode, "Button_Update", "Button_Insert")
            If Button_Insert.Pressed Then
                PressedButton = "Button_Insert"
            ElseIf Button_Update.Pressed Then
                PressedButton = "Button_Update"
            ElseIf Button_Delete.Pressed Then
                PressedButton = "Button_Delete"
            ElseIf Button_Cancel.Pressed Then
                PressedButton = "Button_Cancel"
            End If
        End If
        Redirect = "AdminPageTemplateSectionList.asp"
        If PressedButton = "Button_Delete" Then
            If NOT Button_Delete.OnClick OR NOT DeleteRow() Then
                Redirect = ""
            End If
        ElseIf PressedButton = "Button_Cancel" Then
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
'End PageTemplateSection Operation Method

'PageTemplateSection InsertRow Method @33-E7CA4394
    Function InsertRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeInsert", Me)
        If NOT InsertAllowed Then InsertRow = False : Exit Function
        DataSource.PageTemplateSectionNickname.Value = PageTemplateSectionNickname.Value
        DataSource.PageTemplateSectionSiteID.Value = PageTemplateSectionSiteID.Value
        DataSource.PageTemplateSectionName.Value = PageTemplateSectionName.Value
        DataSource.PageTemplateSectionDescription.Value = PageTemplateSectionDescription.Value
        DataSource.Insert(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInsert", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        InsertRow = (Errors.Count = 0)
    End Function
'End PageTemplateSection InsertRow Method

'PageTemplateSection UpdateRow Method @33-23901DBF
    Function UpdateRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUpdate", Me)
        If NOT UpdateAllowed Then UpdateRow = False : Exit Function
        DataSource.PageTemplateSectionNickname.Value = PageTemplateSectionNickname.Value
        DataSource.PageTemplateSectionSiteID.Value = PageTemplateSectionSiteID.Value
        DataSource.PageTemplateSectionName.Value = PageTemplateSectionName.Value
        DataSource.PageTemplateSectionDescription.Value = PageTemplateSectionDescription.Value
        DataSource.Update(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterUpdate", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        UpdateRow = (Errors.Count = 0)
    End Function
'End PageTemplateSection UpdateRow Method

'PageTemplateSection DeleteRow Method @33-D5C1DF24
    Function DeleteRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeDelete", Me)
        If NOT DeleteAllowed Then DeleteRow = False : Exit Function
        DataSource.Delete(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterDelete", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        DeleteRow = (Errors.Count = 0)
    End Function
'End PageTemplateSection DeleteRow Method

'PageTemplateSection Show Method @33-465F21DA
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "PageTemplateSection" & IIf(EditMode, ":Edit", ""))
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
            Array(PageTemplateSectionNickname, PageTemplateSectionSiteID, PageTemplateSectionName, PageTemplateSectionDescription, Button_Insert, Button_Update, Button_Delete, Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        PageTemplateSectionNickname.Value = Recordset.Fields("PageTemplateSectionNickname")
                        PageTemplateSectionSiteID.Value = Recordset.Fields("PageTemplateSectionSiteID")
                        PageTemplateSectionName.Value = Recordset.Fields("PageTemplateSectionName")
                        PageTemplateSectionDescription.Value = Recordset.Fields("PageTemplateSectionDescription")
                    End If
                Else
                    EditMode = False
                End If
            End If
        End If
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors PageTemplateSectionNickname.Errors
            Errors.AddErrors PageTemplateSectionSiteID.Errors
            Errors.AddErrors PageTemplateSectionName.Errors
            Errors.AddErrors PageTemplateSectionDescription.Errors
            Errors.AddErrors DataSource.Errors
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString()
                .Parse False
            End With
        End If
        TemplateBlock.Variable("Action") = HTMLFormAction
        Button_Insert.Visible = NOT EditMode AND InsertAllowed
        Button_Update.Visible = EditMode AND UpdateAllowed
        Button_Delete.Visible = EditMode AND DeleteAllowed

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Visible Then Controls.Show
    End Sub
'End PageTemplateSection Show Method

End Class 'End PageTemplateSection Class @33-A61BA892

Class clsPageTemplateSectionDataSource 'PageTemplateSectionDataSource Class @33-726CDD35

'DataSource Variables @33-2EB84CC3
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
    Public PageTemplateSectionNickname
    Public PageTemplateSectionSiteID
    Public PageTemplateSectionName
    Public PageTemplateSectionDescription
'End DataSource Variables

'DataSource Class_Initialize Event @33-CD8ABDB3
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set PageTemplateSectionNickname = CCCreateField("PageTemplateSectionNickname", "PageTemplateSectionNickname", ccsText, Empty, Recordset)
        Set PageTemplateSectionSiteID = CCCreateField("PageTemplateSectionSiteID", "PageTemplateSectionSiteID", ccsInteger, Empty, Recordset)
        Set PageTemplateSectionName = CCCreateField("PageTemplateSectionName", "PageTemplateSectionName", ccsText, Empty, Recordset)
        Set PageTemplateSectionDescription = CCCreateField("PageTemplateSectionDescription", "PageTemplateSectionDescription", ccsMemo, Empty, Recordset)
        Fields.AddFields Array(PageTemplateSectionNickname, PageTemplateSectionSiteID, PageTemplateSectionName, PageTemplateSectionDescription)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1  *  " & vbLf & _
        "FROM PageTemplateSection {SQL_Where} {SQL_OrderBy}"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @33-91CDE726
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urlPageTemplateSectionID", ccsInteger, Empty, Empty, Empty, False
            AllParamsSet = .AllParamsSet
            .Criterion(1) = .Operation(opEqual, False, "[PageTemplateSectionID]", .getParamByID(1))
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

'Open Method @33-48A2DA7D
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

'DataSource Class_Terminate Event @33-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

'Delete Method @33-5A0A14CA
    Sub Delete(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildDelete", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        BuildTableWhere
        If NOT AllParamsSet Then
            Errors.AddError(CCSLocales.GetText("CCS_CustomOperationError_MissingParameters", Empty))
        End If
        Cmd.SQL = "DELETE FROM [PageTemplateSection]" & IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteDelete", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteDelete", Me)
        End If
    End Sub
'End Delete Method

'Update Method @33-E299D2FA
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
        Cmd.SQL = "UPDATE [PageTemplateSection] SET " & _
            "[PageTemplateSectionNickname]=" & Connection.ToSQL(PageTemplateSectionNickname, PageTemplateSectionNickname.DataType) & ", " & _
            "[PageTemplateSectionSiteID]=" & Connection.ToSQL(PageTemplateSectionSiteID, PageTemplateSectionSiteID.DataType) & ", " & _
            "[PageTemplateSectionName]=" & Connection.ToSQL(PageTemplateSectionName, PageTemplateSectionName.DataType) & ", " & _
            "[PageTemplateSectionDescription]=?" & _
            IIf(Len(Where) > 0, " WHERE " & Where, "")
        Cmd.CommandParameters = Array( _
            Array("[PageTemplateSectionDescription]", adLongVarChar, adParamInput, 2147483647, PageTemplateSectionDescription.Value))
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

'Insert Method @33-95C4CECA
    Sub Insert(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildInsert", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Cmd.Prepared = True
        Cmd.SQL = "INSERT INTO [PageTemplateSection] (" & _
            "[PageTemplateSectionNickname], " & _
            "[PageTemplateSectionSiteID], " & _
            "[PageTemplateSectionName], " & _
            "[PageTemplateSectionDescription]" & _
        ") VALUES (" & _
            Connection.ToSQL(PageTemplateSectionNickname, PageTemplateSectionNickname.DataType) & ", " & _
            Connection.ToSQL(PageTemplateSectionSiteID, PageTemplateSectionSiteID.DataType) & ", " & _
            Connection.ToSQL(PageTemplateSectionName, PageTemplateSectionName.DataType) & ", " & _
            "?" & _
        ")"
        Cmd.CommandParameters = Array( _
            Array("[PageTemplateSectionDescription]", adLongVarChar, adParamInput,2147483647, PageTemplateSectionDescription.Value) _
        )
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteInsert", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteInsert", Me)
        End If
    End Sub
'End Insert Method

End Class 'End PageTemplateSectionDataSource Class @33-A61BA892


%>
