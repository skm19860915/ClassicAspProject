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

'Initialize Page @1-2598276F
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
Dim Subscription
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminSubscriptionEdit.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminSubscriptionEdit.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Initialize Objects @1-2A339093
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set Subscription = new clsRecordSubscription
Menu.Value = DHTMLMenu

Subscription.Initialize DBSystem

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-38DB9A17
Subscription.Operation
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

'Show Page @1-562CE8D2
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, Subscription))
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

'UnloadPage Sub @1-47611DA7
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set Subscription = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordSubscription 'Subscription Class @33-94675801

'Subscription Variables @33-7EE76AEC

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
    Dim SubscriptionType
    Dim PageTemplateSiteID
    Dim SubscriptionParentID
    Dim SubscriptionName
    Dim Button_Insert
    Dim Button_Update
    Dim Button_Delete
    Dim Button_Cancel
'End Subscription Variables

'Subscription Class_Initialize Event @33-A3EC4545
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsSubscriptionDataSource
        Set Command = New clsCommand
        InsertAllowed = True
        UpdateAllowed = True
        DeleteAllowed = True
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "Subscription")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "Subscription"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set SubscriptionType = CCCreateList(ccsListBox, "SubscriptionType", "Type", ccsText, CCGetRequestParam("SubscriptionType", Method), Empty)
        Set SubscriptionType.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("System", "User"), _
            Array("System", "User")))
        Set PageTemplateSiteID = CCCreateControl(ccsHidden, "PageTemplateSiteID", Empty, ccsInteger, Empty, CCGetRequestParam("PageTemplateSiteID", Method))
        Set SubscriptionParentID = CCCreateList(ccsListBox, "SubscriptionParentID", "Parent", ccsInteger, CCGetRequestParam("SubscriptionParentID", Method), Empty)
        SubscriptionParentID.BoundColumn = "SubscriptionID"
        SubscriptionParentID.TextColumn = "SubscriptionName"
        Set SubscriptionParentID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM Subscription {SQL_Where} {SQL_OrderBy}", "", ""))
        With SubscriptionParentID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("sesSiteID") = Session("SiteID")
            .ParameterSources("urlSubscriptionID") = CCGetRequestParam("SubscriptionID", ccsGET)
            .AddParameter 1, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .AddParameter 2, "urlSubscriptionID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opEqual, False, "[SubscriptionSiteID]", .getParamByID(1))
            .Criterion(2) = .Operation(opNotEqual, False, "[SubscriptionID]", .getParamByID(2))
            .AssembledWhere = .opAND(False, .Criterion(1), .Criterion(2))
        End With
        SubscriptionParentID.DataSource.Where = SubscriptionParentID.DataSource.WhereParameters.AssembledWhere
        Set SubscriptionName = CCCreateControl(ccsTextBox, "SubscriptionName", "Name", ccsText, Empty, CCGetRequestParam("SubscriptionName", Method))
        SubscriptionName.Required = True
        Set Button_Insert = CCCreateButton("Button_Insert", Method)
        Set Button_Update = CCCreateButton("Button_Update", Method)
        Set Button_Delete = CCCreateButton("Button_Delete", Method)
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(SubscriptionType, PageTemplateSiteID, SubscriptionParentID, SubscriptionName)
        If Not FormSubmitted Then
            If IsEmpty(PageTemplateSiteID.Value) Then _
                PageTemplateSiteID.Value = Session("SiteID")
        End If
    End Sub
'End Subscription Class_Initialize Event

'Subscription Initialize Method @33-5F400A16
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("urlSubscriptionID") = CCGetRequestParam("SubscriptionID", ccsGET)
        End With
    End Sub
'End Subscription Initialize Method

'Subscription Class_Terminate Event @33-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End Subscription Class_Terminate Event

'Subscription Validate Method @33-B9D513CF
    Function Validate()
        Dim Validation
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End Subscription Validate Method

'Subscription Operation Method @33-A217B513
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
        Redirect = "AdminSubscriptionList.asp"
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
'End Subscription Operation Method

'Subscription InsertRow Method @33-FF4C6589
    Function InsertRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeInsert", Me)
        If NOT InsertAllowed Then InsertRow = False : Exit Function
        DataSource.SubscriptionType.Value = SubscriptionType.Value
        DataSource.PageTemplateSiteID.Value = PageTemplateSiteID.Value
        DataSource.SubscriptionParentID.Value = SubscriptionParentID.Value
        DataSource.SubscriptionName.Value = SubscriptionName.Value
        DataSource.Insert(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInsert", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        InsertRow = (Errors.Count = 0)
    End Function
'End Subscription InsertRow Method

'Subscription UpdateRow Method @33-D6A8C82B
    Function UpdateRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUpdate", Me)
        If NOT UpdateAllowed Then UpdateRow = False : Exit Function
        DataSource.SubscriptionType.Value = SubscriptionType.Value
        DataSource.PageTemplateSiteID.Value = PageTemplateSiteID.Value
        DataSource.SubscriptionParentID.Value = SubscriptionParentID.Value
        DataSource.SubscriptionName.Value = SubscriptionName.Value
        DataSource.Update(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterUpdate", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        UpdateRow = (Errors.Count = 0)
    End Function
'End Subscription UpdateRow Method

'Subscription DeleteRow Method @33-D5C1DF24
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
'End Subscription DeleteRow Method

'Subscription Show Method @33-B34A573F
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "Subscription" & IIf(EditMode, ":Edit", ""))
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
            Array(SubscriptionType, PageTemplateSiteID, SubscriptionParentID, SubscriptionName, Button_Insert, Button_Update, Button_Delete, Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        SubscriptionType.Value = Recordset.Fields("SubscriptionType")
                        PageTemplateSiteID.Value = Recordset.Fields("PageTemplateSiteID")
                        SubscriptionParentID.Value = Recordset.Fields("SubscriptionParentID")
                        SubscriptionName.Value = Recordset.Fields("SubscriptionName")
                    End If
                Else
                    EditMode = False
                End If
            End If
        End If
        If Not FormSubmitted Then
        End If
        If FormSubmitted Then
            Errors.AddErrors SubscriptionType.Errors
            Errors.AddErrors PageTemplateSiteID.Errors
            Errors.AddErrors SubscriptionParentID.Errors
            Errors.AddErrors SubscriptionName.Errors
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
'End Subscription Show Method

End Class 'End Subscription Class @33-A61BA892

Class clsSubscriptionDataSource 'SubscriptionDataSource Class @33-1E6C1D58

'DataSource Variables @33-8F3A02F6
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
    Public SubscriptionType
    Public PageTemplateSiteID
    Public SubscriptionParentID
    Public SubscriptionName
'End DataSource Variables

'DataSource Class_Initialize Event @33-0F50BE5F
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set SubscriptionType = CCCreateField("SubscriptionType", "SubscriptionType", ccsText, Empty, Recordset)
        Set PageTemplateSiteID = CCCreateField("PageTemplateSiteID", "SubscriptionSiteID", ccsInteger, Empty, Recordset)
        Set SubscriptionParentID = CCCreateField("SubscriptionParentID", "SubscriptionParentID", ccsInteger, Empty, Recordset)
        Set SubscriptionName = CCCreateField("SubscriptionName", "SubscriptionName", ccsText, Empty, Recordset)
        Fields.AddFields Array(SubscriptionType, PageTemplateSiteID, SubscriptionParentID, SubscriptionName)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1  *  " & vbLf & _
        "FROM Subscription {SQL_Where} {SQL_OrderBy}"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @33-065CFFAC
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urlSubscriptionID", ccsInteger, Empty, Empty, Empty, False
            AllParamsSet = .AllParamsSet
            .Criterion(1) = .Operation(opEqual, False, "[SubscriptionID]", .getParamByID(1))
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

'Delete Method @33-DFC0176E
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
        Cmd.SQL = "DELETE FROM [Subscription]" & IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteDelete", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteDelete", Me)
        End If
    End Sub
'End Delete Method

'Update Method @33-3413BEB7
    Sub Update(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildUpdate", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        BuildTableWhere
        If NOT AllParamsSet Then
            Errors.AddError(CCSLocales.GetText("CCS_CustomOperationError_MissingParameters", Empty))
        End If
        Cmd.SQL = "UPDATE [Subscription] SET " & _
            "[SubscriptionType]=" & Connection.ToSQL(SubscriptionType, SubscriptionType.DataType) & ", " & _
            "[SubscriptionSiteID]=" & Connection.ToSQL(PageTemplateSiteID, PageTemplateSiteID.DataType) & ", " & _
            "[SubscriptionParentID]=" & Connection.ToSQL(SubscriptionParentID, SubscriptionParentID.DataType) & ", " & _
            "[SubscriptionName]=" & Connection.ToSQL(SubscriptionName, SubscriptionName.DataType) & _
            IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

'Insert Method @33-0CD8010F
    Sub Insert(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildInsert", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Cmd.SQL = "INSERT INTO [Subscription] (" & _
            "[SubscriptionType], " & _
            "[SubscriptionSiteID], " & _
            "[SubscriptionParentID], " & _
            "[SubscriptionName]" & _
        ") VALUES (" & _
            Connection.ToSQL(SubscriptionType, SubscriptionType.DataType) & ", " & _
            Connection.ToSQL(PageTemplateSiteID, PageTemplateSiteID.DataType) & ", " & _
            Connection.ToSQL(SubscriptionParentID, SubscriptionParentID.DataType) & ", " & _
            Connection.ToSQL(SubscriptionName, SubscriptionName.DataType) & _
        ")"
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteInsert", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteInsert", Me)
        End If
    End Sub
'End Insert Method

End Class 'End SubscriptionDataSource Class @33-A61BA892


%>
