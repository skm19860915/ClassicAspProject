<%@ CodePage=1252 %>
<%
'Include Common Files @1-C42A9701
%>
<!-- #INCLUDE VIRTUAL="/Common.asp"-->
<!-- #INCLUDE VIRTUAL="/Cache.asp" -->
<!-- #INCLUDE VIRTUAL="/Template.asp" -->
<!-- #INCLUDE VIRTUAL="/Sorter.asp" -->
<!-- #INCLUDE VIRTUAL="/Navigator.asp" -->
<!-- #INCLUDE VIRTUAL="/generatecode.asp" -->
<%
'End Include Common Files

'Initialize Page @1-F56A3C46
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
Dim TrackingLinkCode
Dim ChildControls

Redirect = ""
TemplateFileName = "AdminTrackingLinkCodeEdit.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
FileName = "AdminTrackingLinkCodeEdit.asp"
PathToRoot = "./"
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Authenticate User @1-6D464615
CCSecurityRedirect "50;40", Empty
'End Authenticate User

'Initialize Objects @1-4A0630C9
Set DBSystem = New clsDBSystem
DBSystem.Open

' Controls
Set Menu = CCCreateControl(ccsLabel, "Menu", Empty, ccsText, Empty, CCGetRequestParam("Menu", ccsGet))
Menu.HTML = True
Set TrackingLinkCode = new clsRecordTrackingLinkCode
Menu.Value = DHTMLMenu

TrackingLinkCode.Initialize DBSystem

' Events
%>
<!-- #INCLUDE VIRTUAL="AdminTrackingLinkCodeEdit_events.asp" -->
<%
BindEvents

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-0F8D0E93
TrackingLinkCode.Operation
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

'Show Page @1-6F298D88
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Menu, TrackingLinkCode))
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

'UnloadPage Sub @1-1D83FB10
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBSystem.State = adStateOpen Then _
        DBSystem.Close
    Set DBSystem = Nothing
    Set CCSEvents = Nothing
    Set TrackingLinkCode = Nothing
End Sub
'End UnloadPage Sub

Class clsRecordTrackingLinkCode 'TrackingLinkCode Class @2-8A464C35

'TrackingLinkCode Variables @2-C9A8F81F

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
    Dim TrackingLinkCodeSiteID
    Dim TrackingLinkCodeTrackingType
    Dim TrackingLinkCodeRedirectToURL
    Dim TrackingLinkCodeTrackingLinkGroupID
    Dim TrackingLinkCodeDescription
    Dim TrackingLinkCode
    Dim TrackingLinkCodeLink
    Dim Button_Insert
    Dim Button_Update
    Dim Button_Delete
    Dim Button_Cancel
'End TrackingLinkCode Variables

'TrackingLinkCode Class_Initialize Event @2-B0D6677B
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set DataSource = New clsTrackingLinkCodeDataSource
        Set Command = New clsCommand
        InsertAllowed = True
        UpdateAllowed = True
        DeleteAllowed = True
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        OperationMode = Split(CCGetFromGet("ccsForm", Empty), ":")
        If UBound(OperationMode) > -1 Then 
            FormSubmitted = (OperationMode(0) = "TrackingLinkCode")
        End If
        If UBound(OperationMode) > 0 Then 
            EditMode = (OperationMode(1) = "Edit")
        End If
        ComponentName = "TrackingLinkCode"
        Method = IIf(FormSubmitted, ccsPost, ccsGet)
        Set TrackingLinkCodeSiteID = CCCreateControl(ccsHidden, "TrackingLinkCodeSiteID", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkCodeSiteID", Method))
        Set TrackingLinkCodeTrackingType = CCCreateList(ccsListBox, "TrackingLinkCodeTrackingType", "Tracking Code Type", ccsText, CCGetRequestParam("TrackingLinkCodeTrackingType", Method), Empty)
        Set TrackingLinkCodeTrackingType.DataSource = CCCreateDataSource(dsListOfValues, Empty, Array( _
            Array("0", "1", "2"), _
            Array("Incoming", "Outgoing", "Internal")))
        TrackingLinkCodeTrackingType.Required = True
        Set TrackingLinkCodeRedirectToURL = CCCreateControl(ccsTextBox, "TrackingLinkCodeRedirectToURL", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkCodeRedirectToURL", Method))
        Set TrackingLinkCodeTrackingLinkGroupID = CCCreateList(ccsListBox, "TrackingLinkCodeTrackingLinkGroupID", "Tracking Link Group", ccsInteger, CCGetRequestParam("TrackingLinkCodeTrackingLinkGroupID", Method), Empty)
        TrackingLinkCodeTrackingLinkGroupID.BoundColumn = "TrackingLinkGroupID"
        TrackingLinkCodeTrackingLinkGroupID.TextColumn = "TrackingLinkGroupName"
        Set TrackingLinkCodeTrackingLinkGroupID.DataSource = CCCreateDataSource(dsTable,DBSystem, Array("SELECT *  " & _
"FROM TrackingLinkGroup {SQL_Where} {SQL_OrderBy}", "", ""))
        With TrackingLinkCodeTrackingLinkGroupID.DataSource.WhereParameters
            Set .ParameterSources = Server.CreateObject("Scripting.Dictionary")
            .ParameterSources("expr13") = 0
            .ParameterSources("sesSiteID") = Session("SiteID")
            .AddParameter 1, "expr13", ccsInteger, Empty, Empty, Empty, False
            .AddParameter 2, "sesSiteID", ccsInteger, Empty, Empty, Empty, False
            .Criterion(1) = .Operation(opGreaterThan, False, "[TrackingLinkGroupID]", .getParamByID(1))
            .Criterion(2) = .Operation(opEqual, False, "[TrackingLinkGroupSiteID]", .getParamByID(2))
            .AssembledWhere = .opAND(False, .Criterion(1), .Criterion(2))
        End With
        TrackingLinkCodeTrackingLinkGroupID.DataSource.Where = TrackingLinkCodeTrackingLinkGroupID.DataSource.WhereParameters.AssembledWhere
        TrackingLinkCodeTrackingLinkGroupID.Required = True
        Set TrackingLinkCodeDescription = CCCreateControl(ccsTextBox, "TrackingLinkCodeDescription", "Description", ccsText, Empty, CCGetRequestParam("TrackingLinkCodeDescription", Method))
        Set TrackingLinkCode = CCCreateControl(ccsTextBox, "TrackingLinkCode", "Tracking Link Code", ccsText, Empty, CCGetRequestParam("TrackingLinkCode", Method))
        TrackingLinkCode.Required = True
        Set TrackingLinkCodeLink = CCCreateControl(ccsLabel, "TrackingLinkCodeLink", Empty, ccsText, Empty, CCGetRequestParam("TrackingLinkCodeLink", Method))
        Set Button_Insert = CCCreateButton("Button_Insert", Method)
        Set Button_Update = CCCreateButton("Button_Update", Method)
        Set Button_Delete = CCCreateButton("Button_Delete", Method)
        Set Button_Cancel = CCCreateButton("Button_Cancel", Method)
        Set ValidatingControls = new clsControls
        ValidatingControls.addControls Array(TrackingLinkCodeSiteID, TrackingLinkCodeTrackingType, TrackingLinkCodeRedirectToURL, TrackingLinkCodeTrackingLinkGroupID, TrackingLinkCodeDescription, TrackingLinkCode)
        If Not FormSubmitted Then
            If IsEmpty(TrackingLinkCodeSiteID.Value) Then _
                TrackingLinkCodeSiteID.Value = Session("SiteID")
            If IsEmpty(TrackingLinkCodeTrackingLinkGroupID.Value) Then _
                TrackingLinkCodeTrackingLinkGroupID.Value = Request.QueryString("TrackingLinkGroupID")
        End If
    End Sub
'End TrackingLinkCode Class_Initialize Event

'TrackingLinkCode Initialize Method @2-8595C5C8
    Sub Initialize(objConnection)

        If NOT Visible Then Exit Sub


        Set DataSource.Connection = objConnection
        With DataSource
            .Parameters("urlTrackingLinkCodeID") = CCGetRequestParam("TrackingLinkCodeID", ccsGET)
        End With
    End Sub
'End TrackingLinkCode Initialize Method

'TrackingLinkCode Class_Terminate Event @2-32B847C9
    Private Sub Class_Terminate()
        Set Errors = Nothing
    End Sub
'End TrackingLinkCode Class_Terminate Event

'TrackingLinkCode Validate Method @2-1324BC3D
    Function Validate()
        Dim Validation
        Dim Where
        If EditMode Then
            DataSource.BuildTableWhere
            If DataSource.Where <> "" Then _
                Where = " AND NOT (" & DataSource.Where & ")"
        End If
        If TrackingLinkCode.Errors.Count = 0 Then _
            If CInt(CCDLookUp("COUNT(*)", " TrackingLinkCode", "[TrackingLinkCode] =" & DBSystem.ToSQL(TrackingLinkCode.Value, TrackingLinkCode.DataType) & Where, DBSystem)) > 0 Then _
                TrackingLinkCode.Errors.addError(CCSLocales.GetText("CCS_UniqueValue", Array(TrackingLinkCode.Caption)))
        ValidatingControls.Validate
        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = ValidatingControls.isValid() And (Errors.Count = 0)
    End Function
'End TrackingLinkCode Validate Method

'TrackingLinkCode Operation Method @2-15395901
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
        Redirect = "AdminTrackingLinkGroupList.asp"
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
'End TrackingLinkCode Operation Method

'TrackingLinkCode InsertRow Method @2-C684816D
    Function InsertRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeInsert", Me)
        If NOT InsertAllowed Then InsertRow = False : Exit Function
        DataSource.TrackingLinkCodeSiteID.Value = TrackingLinkCodeSiteID.Value
        DataSource.TrackingLinkCodeTrackingType.Value = TrackingLinkCodeTrackingType.Value
        DataSource.TrackingLinkCodeRedirectToURL.Value = TrackingLinkCodeRedirectToURL.Value
        DataSource.TrackingLinkCodeTrackingLinkGroupID.Value = TrackingLinkCodeTrackingLinkGroupID.Value
        DataSource.TrackingLinkCodeDescription.Value = TrackingLinkCodeDescription.Value
        DataSource.TrackingLinkCode.Value = TrackingLinkCode.Value
        DataSource.Insert(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInsert", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        InsertRow = (Errors.Count = 0)
    End Function
'End TrackingLinkCode InsertRow Method

'TrackingLinkCode UpdateRow Method @2-FD1CD2A8
    Function UpdateRow()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUpdate", Me)
        If NOT UpdateAllowed Then UpdateRow = False : Exit Function
        DataSource.TrackingLinkCodeSiteID.Value = TrackingLinkCodeSiteID.Value
        DataSource.TrackingLinkCodeTrackingType.Value = TrackingLinkCodeTrackingType.Value
        DataSource.TrackingLinkCodeRedirectToURL.Value = TrackingLinkCodeRedirectToURL.Value
        DataSource.TrackingLinkCodeTrackingLinkGroupID.Value = TrackingLinkCodeTrackingLinkGroupID.Value
        DataSource.TrackingLinkCodeDescription.Value = TrackingLinkCodeDescription.Value
        DataSource.TrackingLinkCode.Value = TrackingLinkCode.Value
        DataSource.Update(Command)


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterUpdate", Me)
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
        End If
        UpdateRow = (Errors.Count = 0)
    End Function
'End TrackingLinkCode UpdateRow Method

'TrackingLinkCode DeleteRow Method @2-D5C1DF24
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
'End TrackingLinkCode DeleteRow Method

'TrackingLinkCode Show Method @2-99C3F0B4
    Sub Show(Tpl)

        If NOT Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        EditMode = Recordset.EditMode(ReadAllowed)
        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "TrackingLinkCode" & IIf(EditMode, ":Edit", ""))
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
            Array(TrackingLinkCodeSiteID,  TrackingLinkCodeTrackingType,  TrackingLinkCodeRedirectToURL,  TrackingLinkCodeTrackingLinkGroupID,  TrackingLinkCodeDescription,  TrackingLinkCode,  TrackingLinkCodeLink, _
                 Button_Insert,  Button_Update,  Button_Delete,  Button_Cancel))
        If EditMode AND ReadAllowed Then
            If Errors.Count = 0 Then
                If Recordset.Errors.Count > 0 Then
                    With TemplateBlock.Block("Error")
                        .Variable("Error") = Recordset.Errors.ToString
                        .Parse False
                    End With
                ElseIf Recordset.CanPopulate() Then
                    If Not FormSubmitted Then
                        TrackingLinkCodeSiteID.Value = Recordset.Fields("TrackingLinkCodeSiteID")
                        TrackingLinkCodeTrackingType.Value = Recordset.Fields("TrackingLinkCodeTrackingType")
                        TrackingLinkCodeRedirectToURL.Value = Recordset.Fields("TrackingLinkCodeRedirectToURL")
                        TrackingLinkCodeTrackingLinkGroupID.Value = Recordset.Fields("TrackingLinkCodeTrackingLinkGroupID")
                        TrackingLinkCodeDescription.Value = Recordset.Fields("TrackingLinkCodeDescription")
                        TrackingLinkCode.Value = Recordset.Fields("TrackingLinkCode")
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
            Errors.AddErrors TrackingLinkCodeSiteID.Errors
            Errors.AddErrors TrackingLinkCodeTrackingType.Errors
            Errors.AddErrors TrackingLinkCodeRedirectToURL.Errors
            Errors.AddErrors TrackingLinkCodeTrackingLinkGroupID.Errors
            Errors.AddErrors TrackingLinkCodeDescription.Errors
            Errors.AddErrors TrackingLinkCode.Errors
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
'End TrackingLinkCode Show Method

End Class 'End TrackingLinkCode Class @2-A61BA892

Class clsTrackingLinkCodeDataSource 'TrackingLinkCodeDataSource Class @2-2803142A

'DataSource Variables @2-E39B92AE
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
    Public TrackingLinkCodeSiteID
    Public TrackingLinkCodeTrackingType
    Public TrackingLinkCodeRedirectToURL
    Public TrackingLinkCodeTrackingLinkGroupID
    Public TrackingLinkCodeDescription
    Public TrackingLinkCode
'End DataSource Variables

'DataSource Class_Initialize Event @2-1DD72469
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set TrackingLinkCodeSiteID = CCCreateField("TrackingLinkCodeSiteID", "TrackingLinkCodeSiteID", ccsText, Empty, Recordset)
        Set TrackingLinkCodeTrackingType = CCCreateField("TrackingLinkCodeTrackingType", "TrackingLinkCodeTrackingType", ccsText, Empty, Recordset)
        Set TrackingLinkCodeRedirectToURL = CCCreateField("TrackingLinkCodeRedirectToURL", "TrackingLinkCodeRedirectToURL", ccsText, Empty, Recordset)
        Set TrackingLinkCodeTrackingLinkGroupID = CCCreateField("TrackingLinkCodeTrackingLinkGroupID", "TrackingLinkCodeTrackingLinkGroupID", ccsInteger, Empty, Recordset)
        Set TrackingLinkCodeDescription = CCCreateField("TrackingLinkCodeDescription", "TrackingLinkCodeDescription", ccsText, Empty, Recordset)
        Set TrackingLinkCode = CCCreateField("TrackingLinkCode", "TrackingLinkCode", ccsText, Empty, Recordset)
        Fields.AddFields Array(TrackingLinkCodeSiteID, TrackingLinkCodeTrackingType, TrackingLinkCodeRedirectToURL, TrackingLinkCodeTrackingLinkGroupID, TrackingLinkCodeDescription, TrackingLinkCode)
        Set Parameters = Server.CreateObject("Scripting.Dictionary")
        Set WhereParameters = Nothing

        SQL = "SELECT TOP 1  *  " & vbLf & _
        "FROM TrackingLinkCode {SQL_Where} {SQL_OrderBy}"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'BuildTableWhere Method @2-30DD6D07
    Public Sub BuildTableWhere()
        Dim WhereParams

        If Not WhereParameters Is Nothing Then _
            Exit Sub
        Set WhereParameters = new clsSQLParameters
        With WhereParameters
            Set .Connection = Connection
            Set .ParameterSources = Parameters
            Set .DataSource = Me
            .AddParameter 1, "urlTrackingLinkCodeID", ccsInteger, Empty, Empty, Empty, False
            AllParamsSet = .AllParamsSet
            .Criterion(1) = .Operation(opEqual, False, "[TrackingLinkCodeID]", .getParamByID(1))
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

'Open Method @2-48A2DA7D
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

'DataSource Class_Terminate Event @2-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

'Delete Method @2-280F191F
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
        Cmd.SQL = "DELETE FROM [TrackingLinkCode]" & IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteDelete", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteDelete", Me)
        End If
    End Sub
'End Delete Method

'Update Method @2-AD601DB5
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
        Cmd.SQL = "UPDATE [TrackingLinkCode] SET " & _
            "[TrackingLinkCodeSiteID]=" & Connection.ToSQL(TrackingLinkCodeSiteID, TrackingLinkCodeSiteID.DataType) & ", " & _
            "[TrackingLinkCodeTrackingType]=" & Connection.ToSQL(TrackingLinkCodeTrackingType, TrackingLinkCodeTrackingType.DataType) & ", " & _
            "[TrackingLinkCodeRedirectToURL]=" & Connection.ToSQL(TrackingLinkCodeRedirectToURL, TrackingLinkCodeRedirectToURL.DataType) & ", " & _
            "[TrackingLinkCodeTrackingLinkGroupID]=" & Connection.ToSQL(TrackingLinkCodeTrackingLinkGroupID, TrackingLinkCodeTrackingLinkGroupID.DataType) & ", " & _
            "[TrackingLinkCodeDescription]=" & Connection.ToSQL(TrackingLinkCodeDescription, TrackingLinkCodeDescription.DataType) & ", " & _
            "[TrackingLinkCode]=" & Connection.ToSQL(TrackingLinkCode, TrackingLinkCode.DataType) & _
            IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

'Insert Method @2-96B7603A
    Sub Insert(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildInsert", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Cmd.SQL = "INSERT INTO [TrackingLinkCode] (" & _
            "[TrackingLinkCodeSiteID], " & _
            "[TrackingLinkCodeTrackingType], " & _
            "[TrackingLinkCodeRedirectToURL], " & _
            "[TrackingLinkCodeTrackingLinkGroupID], " & _
            "[TrackingLinkCodeDescription], " & _
            "[TrackingLinkCode]" & _
        ") VALUES (" & _
            Connection.ToSQL(TrackingLinkCodeSiteID, TrackingLinkCodeSiteID.DataType) & ", " & _
            Connection.ToSQL(TrackingLinkCodeTrackingType, TrackingLinkCodeTrackingType.DataType) & ", " & _
            Connection.ToSQL(TrackingLinkCodeRedirectToURL, TrackingLinkCodeRedirectToURL.DataType) & ", " & _
            Connection.ToSQL(TrackingLinkCodeTrackingLinkGroupID, TrackingLinkCodeTrackingLinkGroupID.DataType) & ", " & _
            Connection.ToSQL(TrackingLinkCodeDescription, TrackingLinkCodeDescription.DataType) & ", " & _
            Connection.ToSQL(TrackingLinkCode, TrackingLinkCode.DataType) & _
        ")"
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteInsert", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteInsert", Me)
        End If
    End Sub
'End Insert Method

End Class 'End TrackingLinkCodeDataSource Class @2-A61BA892


%>
