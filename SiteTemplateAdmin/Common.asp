<%
Option Explicit

'Include Files @0-7F03F314
%>
<!-- #INCLUDE FILE="Adovbs.asp" -->
<!-- #INCLUDE FILE="Classes.asp" -->
<!-- #INCLUDE FILE="i_Menu.asp" -->
<%
'End Include Files

'Script Engine Version Check @0-02EA3A85
If ScriptEngineMajorVersion < 5 Then
  Response.Write "Sorry. This program requires VBScript 5.1 to run.<br>You may upgrade your VBScript at http://msdn.microsoft.com/downloads/list/webdev.asp?frame=true."
  Response.End
Else
  If ScriptEngineMajorVersion & ":" & ScriptEngineMinorVersion = "5:0" Then
    Response.Write "Due to a bug in VBScript 5.0, this program would crash your server. See http://support.microsoft.com/default.aspx?scid=kb;EN-US;q240811.<br>" & _
      "Upgrade your VBScript at http://msdn.microsoft.com/downloads/list/webdev.asp?frame=true."
    Response.End
  End If
End If
'End Script Engine Version Check

'Initialize Common Variables @0-2413C994
<!-- Dim ServerURL : ServerURL = "http://staback.promosonline.com/" -->
Dim ServerURL : ServerURL = "http://localhost/"
Dim SecureURL : SecureURL = ""
Dim InputCodePage : InputCodePage = Session.CodePage
Dim CCSDateConstants
Dim TemplatesRepository
Dim EventCaller
Dim ParentPage
Dim DefaultDateFormat
Dim DefaultBooleanFormat
Dim IsMutipartEncoding
Dim objUpload
Dim UploadedFilesCount
Dim CCSConverter
Dim CCSLocales
Dim CCSStyle
IsMutipartEncoding = False
If InStr(Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data") > 0 And CCGetFromGet("ccsForm", "") <> "" Then
  Set objUpload = new clsUploadControl
  UploadedFilesCount = objUpload.FilesCount
  IsMutipartEncoding = True
End If
Set CCSLocales = New clsLocales
With CCSLocales
  .AppPrefix = "SiteTemplate_Locales_"
  .PathRes = Server.MapPath("/")
  .Locales.Add "en", "US"
  CCLoadStaticTranslation
  .SelectLocale "en", Array("locale", Empty, "locale", "lang", Empty), 365
  .Locale.Charset = "windows-1252"
  .Locale.CodePage = 1252
End With
CCSStyle = "Blueprint"
Set TemplatesRepository = New clsCache_FileSystem
DefaultDateFormat = IIF(CCSLocales.Locale.OverrideDateFormats, CCSLocales.Locale.ShortDate, Array("ShortDate"))
DefaultBooleanFormat = CCSLocales.Locale.BooleanFormat
Set CCSConverter = New clsConverter
CCSConverter.DateFormat = DefaultDateFormat
CCSConverter.BooleanFormat = DefaultBooleanFormat


Set CCSDateConstants = New clsCCSDateConstants

Class clsCCSDateConstants

  Public Weekdays
  Public ShortWeekdays
  Public Months
  Public ShortMonths
  Public DateMasks

  Private Sub Class_Initialize()
    ShortWeekdays = CCSLocales.Locale.WeekdayShortNames
    Weekdays = CCSLocales.Locale.WeekdayNames
    ShortMonths =  CCSLocales.Locale.MonthShortNames
    Months = CCSLocales.Locale.MonthNames
    Set DateMasks = CreateObject("Scripting.Dictionary")
    DateMasks("d") = 0
    DateMasks("dd") = 2
    DateMasks("ddd") = 0
    DateMasks("dddd") = 0
    DateMasks("m") = 0
    DateMasks("mm") = 2
    DateMasks("mmm") = 3
    DateMasks("mmmm") = 0
    DateMasks("yy") = 2
    DateMasks("yyyy") = 4
    DateMasks("h") = 0
    DateMasks("hh") = 2
    DateMasks("H") = 0
    DateMasks("HH") = 2
    DateMasks("n") = 0
    DateMasks("nn") = 2
    DateMasks("s") = 0
    DateMasks("ss") = 2
    DateMasks("am/pm") = 2
    DateMasks("AM/PM") = 2
    DateMasks("A/P") = 1
    DateMasks("a/p") = 1
    DateMasks("w") = 0
    DateMasks("q") = 0
    DateMasks("S") = 0
    DateMasks("tt") = 2
    DateMasks("wi") = 2
  End Sub

  Private Sub Class_Terminate()
    Set DateMasks = Nothing
  End Sub

End Class

Const ccsInteger = 1
Const ccsFloat = 2
Const ccsText = 3
Const ccsDate = 4
Const ccsBoolean = 5
Const ccsMemo = 6
Const ccsSingle = 7
Const ccsGet = 1
Const ccsPost = 2

Const calYear = 0
Const calQuarter = 1
Const cal3Month = 2
Const calMonth = 3
Const calWeek = 4
Const calDay = 5
'End Initialize Common Variables

'System Connection Class @-1882E203
Class clsDBSystem

    Public ConnectionString
    Public User
    Public Password
    Public LastSQL
    Public Errors
    Public Converter
    Public Database

    Private mDateFormat
    Private mBooleanFormat
    Private objConnection
    Private blnState

    

    Private Sub Class_Initialize()
        ConnectionString = "Provider=SQLOLEDB; Data Source = (local); Initial Catalog = SiteTemplate; User Id = sa; Password=123"
        User = "sa"
        Password = "123"
        Set Converter = New clsConverter
        Converter.DateFormat = Array("yyyy", "-", "mm", "-", "dd", " ", "HH", ":", "nn", ":", "ss")
        Converter.BooleanFormat = Array(1, 0, Empty)
        Set objConnection = Server.CreateObject("ADODB.Connection")
        Database = "MSSQLServer"
        Set Errors = New clsErrors
    End Sub

    Public Property Get DateFormat()
      DateFormat = Converter.DateFormat
    End Property

    Public Property Get BooleanFormat()
      BooleanFormat = Converter.BooleanFormat
    End Property

    Sub Open()
        On Error Resume Next
        objConnection.Errors.Clear
        objConnection.Open ConnectionString, User, Password
        If Err.Number <> 0 then
            Response.Write "<div><h2>Unable to establish connection to database.</h2>"
            Response.Write "<ul><li><b>Error information:</b><br>"
            Response.Write Err.Source & " (0x" & Hex(Err.Number) & ")<br>"
            Response.Write Err.Description & "</li>"
            If Err.Number = -2147467259 then _
            Response.Write "<li><b>Other possible cause of this problem:</b><br>The database cannot be opened, most likely due to incorrect connection settings or insufficient security set on your database folder or file. <br>For more details please refer to <a href='http://support.microsoft.com/default.aspx?scid=kb;en-us;Q306518'>http://support.microsoft.com/default.aspx?scid=kb;en-us;Q306518</a></li>"
            Response.Write "</ul></div>"
            Response.End
        End If
    End Sub

    Sub Close()
        objConnection.Close
    End Sub

    Function Execute(varCMD)
        Dim ErrorMessage, objResult
        Errors.Clear
        Set objResult = Server.CreateObject("ADODB.Recordset")
        objResult.CursorType = adOpenForwardOnly
        objResult.LockType = adLockReadOnly
        If TypeName(varCMD) = "Command" Then
            Set varCMD.ActiveConnection = objConnection
            Set objResult.Source = varCMD
            LastSQL = varCMD.CommandText
        Else
            Set objResult.ActiveConnection = objConnection
            objResult.Source = varCMD
            LastSQL = varCMD
        End If
        On Error Resume Next
        objResult.Open
        Errors.AddError CCProcessError(objConnection)
        On Error Goto 0
        Set Execute = objResult
    End Function

    Property Get Connection()
        Set Connection = objConnection
    End Property

    Property Get State()
        State = objConnection.State
    End Property

    Function ToSQL(Value, ValueType)
        Dim mValue
        Dim needEscape : needEscape = True
        If TypeName(Value) = "clsSQLParameter" or TypeName(Value) = "clsField" Then 
            mValue = Value.SQLText
             needEscape = False
        Else 
            mValue = Value
        End If
        If CStr(mValue) = "" OR IsEmpty(mValue) Then
            ToSQL = "Null"
        Else
            Select Case ValueType
                Case ccsDate
                    If VarType(mValue)=vbDate And TypeName(Value) <> "clsSQLParameter" Then _
                        mValue = CCFormatDate(mValue, DateFormat)
                Case ccsBoolean
                    If VarType(mValue)=vbBoolean And TypeName(Value) <> "clsSQLParameter" Then _
                        mValue= CCFormatBoolean(mValue, BooleanFormat)
            End Select

            If ValueType = ccsInteger or ValueType = ccsFloat or ValueType = ccsSingle Then
                ToSQL = Replace(mValue, ",", ".")
            ElseIf ValueType = ccsDate Then
                ToSQL = "'" & mValue & "'"
            Else
                If needEscape And CStr(mValue) <> "" Then mValue = EscapeChars(mValue)
                ToSQL = "N'" & mValue & "'"
            End If
        End If
    End Function

    Function ToLikeCriteria(Value, CriteriaType)
        Select Case CriteriaType
            Case opBeginsWith
                ToLikeCriteria =  " like N'" & Value & "%'"
            Case opNotBeginsWith
                ToLikeCriteria =  " not like N'" & Value & "%'"
            Case opEndsWith
                ToLikeCriteria =  " like N'%" & Value & "'"
            Case opNotEndsWith
                ToLikeCriteria =  " not like N'%" & Value & "'"
            Case opContains
                ToLikeCriteria =  " like N'%" & Value & "%'"
            Case opNotContains
                ToLikeCriteria =  " not like N'%" & Value & "%'"
        End Select
    End Function

    Function EscapeChars(Value)
        EscapeChars = Replace(Value, "'", "''")
    End Function

End Class
'End System Connection Class

'IIf @0-E12349E2
Function IIf(Expression, TrueResult, FalseResult)
  If CBool(Expression) Then
    If IsObject(TrueResult) Then _
      Set IIf = TrueResult _
    Else _
      IIf = TrueResult
  Else
    If IsObject(FalseResult) Then _
      Set IIf = FalseResult _
    Else _
      IIf = FalseResult
  End If
End Function
'End IIf

'Print @0-065FC167
Sub Print(Value)
  Response.Write CStr(Value)
End Sub
'End Print

'CCRaiseEvent @0-5BA4885B
Function CCRaiseEvent(Events, EventName, Caller)
  Dim Result
  Dim EC : Set EC = New clsEventCaller
  If Events.Exists(EventName) Then
    Set EventCaller = Caller
    Set EC.EventRef = Events(EventName)
    Result = EC.Invoke(Caller)
  End If
  Set EventCaller = Nothing
  If VarType(Result) = vbEmpty Then _
    Result = True
  CCRaiseEvent = Result
End Function
'End CCRaiseEvent

'CCFormatError @0-BB43933D
Function CCFormatError(Title, Errors)
  Dim Result, i
  Result = "<p>Form: " & CCToHTML(Title) & "<br>"
  For i = 0 To Errors.Count - 1
    Result = Result & "Error: " & Replace(CCToHTML(Replace(Errors.ErrorByNumber(i), "<br>", vbCrLf)), vbCrLf, "<br>")
    If i < Errors.Count - 1 Then Result = Result & "<br>"
  Next
  Result = Result & "</p>"
  CCFormatError = Result
End Function
'End CCFormatError

'CCOpenRS @0-B04DFBEE
Function CCOpenRS(RecordSet, SQL, Connection, ShowError)
  Dim ErrorMessage, Result
  Result = Empty
  Set RecordSet = Server.CreateObject("ADODB.Recordset")
  On Error Resume Next
  RecordSet.Open SQL, Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
  ErrorMessage = CCProcessError(Connection)
  If NOT IsEmpty(ErrorMessage) Then
    If ShowError Then
      Result = "SQL: " & CommandObject.CommandText & "<br>" & "Error: " & ErrorMessage & "<br>"
    Else
      Result = "Database error.<br>"
    End If
  End If
  On Error Goto 0
  CCOpenRS = Result

End Function
'End CCOpenRS

'CCOpenRSFromCmd @0-25A93885
Function CCOpenRSFromCmd(RecordSet, CommandObject, ShowError)

  Dim ErrorMessage, Result
  Result = Empty
  Set RecordSet = Server.CreateObject("ADODB.Recordset")
  On Error Resume Next
  RecordSet.CursorType = adOpenForwardOnly
  RecordSet.LockType = adLockReadOnly
  RecordSet.Open CommandObject
  ErrorMessage = CCProcessError(CommandObject.ActiveConnection)
  If NOT IsEmpty(ErrorMessage) Then
    If ShowError Then
      Result = "SQL: " & CommandObject.CommandText & "<br>" & "Error: " & ErrorMessage & "<br>"
    Else
      Result = "Database error.<br>"
    End If
  End If
  On Error Goto 0
  CCOpenRSFromCmd = Result

End Function
'End CCOpenRSFromCmd

'CCExecCmd @0-08BE568E
Function CCExecCmd(CommandObject, ShowError)
  Dim ErrorMessage, Result
  Result = Empty
  On Error Resume Next
  CommandObject.Execute
  ErrorMessage = CCProcessError(CommandObject.ActiveConnection)
  If NOT IsEmpty(ErrorMessage) Then 
    If ShowError Then
      Result = "SQL: " & CommandObject.CommandText & "<br>" & "Error: " & ErrorMessage & "<br>"
    Else
      Result = "Database error.<br>"
    End If
  End If
  On Error Goto 0
  CCExecCmd = Result
End Function
'End CCExecCmd

'CCExecSQL @0-1CBAE603
Function CCExecSQL(SQL, Connection, ShowError)
  Dim ErrorMessage, Result
  Result = Empty
  On Error Resume Next
  Connection.Execute(SQL)
  ErrorMessage = CCProcessError(Connection)
  If NOT IsEmpty(ErrorMessage) Then
    If ShowError Then
      Result = "SQL: " & SQL & "<br>" & "Error: " & ErrorMessage & "<br>"
    Else
      Result = "Database error.<br>"
    End If
  End If
  On Error Goto 0
  CCExecSQL = Result
End Function
'End CCExecSQL

'CCToHTML @0-44D2E9F4
Function CCToHTML(Value)
  If IsNull(Value) Then Value = ""
  CCToHTML = Server.HTMLEncode(Value)
End Function
'End CCToHTML

'CCToURL @0-23A93674
Function CCToURL(Value)
  If IsNull(Value) Then Value = ""
  CCToURL = Server.URLEncode(Value)
End Function
'End CCToURL

'CCEscapeLOV @0-B9505CBC
Function CCEscapeLOV(Value)
  CCEscapeLOV = Replace(Replace(CStr(Value), "\", "\\"), ";", "\;")
End Function
'End CCEscapeLOV

'CCUnEscapeLOV @0-4C1E08FE
Function CCUnEscapeLOV(Value)
  CCUnEscapeLOV = Replace(Replace(CStr(Value), "\;", ";"), "\\", "\")
End Function
'End CCUnEscapeLOV

'CCGetValueHTML @0-30C69AED
Function CCGetValueHTML(RecordSet, FieldName)
  CCGetValueHTML = CCToHTML(CCGetValue(RecordSet, FieldName))
End Function
'End CCGetValueHTML

'CCGetValue @0-C5915067
Function CCGetValue(RecordSet, FieldName)
  Dim Result
  On Error Resume Next
  If RecordSet Is Nothing Then
    CCGetValue = Empty
  ElseIf (NOT RecordSet.EOF) AND (FieldName <> "") Then
    Result = RecordSet(FieldName)
    If IsNull(Result) Then _
      Result = Empty
    CCGetValue = Result
  Else
    CCGetValue = Empty
  End If
  On Error Goto 0
End Function
'End CCGetValue

'CCGetDate @0-4102C01B
Function CCGetDate(RecordSet, FieldName, arrDateFormat)
  Dim Result  
  Result = CCGetValue(RecordSet, FieldName)
  If Not IsEmpty(arrDateFormat) Then 
    If Not (VarType(Result) = vbDate OR VarType(Result) = vbEmpty) Then _
      If CCValidateDate(Result, arrDateFormat) Then _
        Result = CCParseDate(Result, arrDateFormat)
  End If
  CCGetDate = Result
End Function
'End CCGetDate

'CCGetBoolean @0-C64EED38
Function CCGetBoolean(RecordSet, FieldName, BooleanFormat)
  Dim Result
  Result = CCGetValue(RecordSet, FieldName)
  CCGetBoolean = CCParseBoolean(Result, BooleanFormat)
End Function
'End CCGetBoolean

'CCGetParam @0-07E4C55C
Function CCGetParam(ParameterName, DefaultValue)
  Dim ParameterValue : ParameterValue = ""

  If IsMutipartEncoding Then
    If Request.QueryString(ParameterName).Count > 0 Then 
      ParameterValue = Request.QueryString(ParameterName)
    Else
      ParameterValue = objUpload.Form(ParameterName)
    End If
    If ParameterValue = "" Then ParameterValue = DefaultValue
  Else
    If Request.QueryString(ParameterName).Count > 0 Then 
      ParameterValue = Request.QueryString(ParameterName)
    ElseIf Request.Form(ParameterName).Count > 0 Then
      ParameterValue = Request.Form(ParameterName)
    Else 
      ParameterValue = DefaultValue
    End If
  End If

  CCGetParam = ParameterValue
End Function
'End CCGetParam

'CCGetFromPost @0-EB8B7999
Function CCGetFromPost(ParameterName, DefaultValue)
  Dim ParameterValue : ParameterValue = Empty

  If IsMutipartEncoding Then
    ParameterValue = objUpload.Form(ParameterName)
    If ParameterValue = "" Then ParameterValue = DefaultValue
  Else
    ParameterValue = Request.Form(ParameterName)
    If IsEmpty(ParameterValue) Then 
      ParameterValue = DefaultValue
    End If
  End If

  CCGetFromPost = ParameterValue
End Function
'End CCGetFromPost

'CCGetFromGet @0-F6BB8115
Function CCGetFromGet(ParameterName, DefaultValue)
  Dim ParameterValue : ParameterValue = Empty
  ParameterValue = Request.QueryString(ParameterName)
  If IsEmpty(ParameterValue) Then _
    ParameterValue = DefaultValue
  CCGetFromGet = ParameterValue
End Function
'End CCGetFromGet

'CCGetCookie @0-25444E9C
Function CCGetCookie(Name)
 Dim q
 For Each q in Request.Cookies
   If (q = Name) Then 
     CCGetCookie = Request.Cookies(q)
     Exit Function
   End If
 Next
 CCGetCookie = Empty
End Function
'End CCGetCookie

'CCToSQL @0-131A8AA9
Function CCToSQL(Value, ValueType)
  If CStr(Value) = "" OR IsEmpty(Value) Then
   CCToSQL = "Null"
  Else
    If ValueType = "Integer" or ValueType = "Float" Then
      CCToSQL = Replace(CDbl(Value), ",", ".")
    ElseIf  ValueType = "Single" Then 
      CCToSQL = Replace(CSng(Value), ",", ".")
    Else 
      CCToSQL = "'" & Replace(Value, "'", "''") & "'"
    End If
  End If
End Function
'End CCToSQL

'CCDLookUp @0-49D86A74
Function CCDLookUp(ColumnName, TableName, Where, Connection)
  Dim RecordSet
  Dim Result
  Dim SQL
  Dim ErrorMessage

  SQL = "SELECT " & ColumnName 
  If Len(CStr(TableName)) > 0 Then SQL = SQL & " FROM "  & TableName 
  If Len(CStr(Where))     > 0 Then SQL = SQL & " WHERE " & Where

  Set RecordSet = Connection.Execute(SQL)
  ErrorMessage = CCProcessError(Connection)
  If NOT IsEmpty(ErrorMessage) Then
    PrintDBError "CCDLookUp function", CCToHTML(SQL), ErrorMessage
  End If
  On Error Goto 0
  Result = CCGetValue(RecordSet, 0)
  CCDLookUp = Result
End Function
'End CCDLookUp

'Min @0-D2DE75DE
Function Min(Value1, Value2)
  Dim result
  If IsEmpty(Value1) Then Min = Value2
  If IsEmpty(Value2) Then Min = Value1
  If Not IsEmpty(Value1) And Not IsEmpty(Value2) Then 
    If Value1 < Value2 Then 
      Min = Value1
    Else
      Min = Value2
    End If
  End If
End Function
'End Min

'Max @0-E097390A
Function Max(Value1, Value2)
  Dim result
  If IsEmpty(Value1) Then Max = Value2
  If IsEmpty(Value2) Then Max = Value1
  If Not IsEmpty(Value1) And Not IsEmpty(Value2) Then 
    If Value1 > Value2 Then 
      Max = Value1
    Else
      Max = Value2
    End If
  End If
End Function
'End Max

'CCGetOriginalFileName @0-3A7CB06E
Function CCGetOriginalFileName(Value)
  If CCRegExpTest(Value, "^\d{14,}\.", True,True) Then
    CCGetOriginalFileName = Mid(Value, InStr(Value,".")+1)   
  Else 
   CCGetOriginalFileName = Value
  End If
End Function
'End CCGetOriginalFileName

'PrintDBError @0-8BC04DA8
Sub PrintDBError(Source, SQL, ErrorMessage)
  Dim CommandText
  Dim SourceText
  Dim ErrorText

  If Source <> "" Then SourceText = "<b>Source:</b> " & Source & "<br>"
  If SQL <> "" Then CommandText = "<b>Command Text:</b> " & SQL & "<br>"
  If ErrorMessage <> "" Then ErrorText = "<b>Error description:</b> " & CCToHTML(ErrorMessage) & "</div>"

  Response.Write "<div style=""background-color: rgb(250, 250, 250); " & _
    "border: solid 1px rgb(200, 200, 200);"">" & SourceText
  Response.Write CommandText & ErrorText
End Sub
'End PrintDBError

'CCGetCheckBoxValue @0-E17ABD19
Function CCGetCheckBoxValue(Value, CheckedValue, UncheckedValue, ValueType)
  If isEmpty(Value) Then
    If UncheckedValue = "" Then
      CCGetCheckBoxValue = "Null"
    Else
      If ValueType = "Integer" or ValueType = "Float" or ValueType = "Single" Then
        CCGetCheckBoxValue = UncheckedValue
      Else
        CCGetCheckBoxValue = "'" & Replace(UncheckedValue, "'", "''") & "'"
      End If
    End If
  Else
    If CheckedValue = "" Then
      CCGetCheckBoxValue = "Null"
    Else
      If ValueType = "Integer" OR ValueType = "Float"  OR ValueType = "Single" Then
        CCGetCheckBoxValue = CheckedValue
      Else
        CCGetCheckBoxValue = "'" & Replace(CheckedValue, "'", "''") & "'"
      End If
    End If
  End If
End Function
'End CCGetCheckBoxValue

'CCGetValFromLOV @0-5041B9C1
Function CCGetValFromLOV(Value, ListOfValues)
  Dim I
  Dim Result : Result = ""
  If (Ubound(ListOfValues) MOD 2) = 1 Then
    For I = 0 To Ubound(ListOfValues) Step 2
      If CStr(Value) = CStr(ListOfValues(I)) Then Result = ListOfValues(I + 1)
    Next
  End If
  CCGetValFromLOV = Result  
End Function
'End CCGetValFromLOV

'CCProcessError @0-A3A2654C
Function CCProcessError(Connection)
  If Connection.Errors.Count > 0 Then
    If TypeName(Connection) = "Connection" Then
      CCProcessError = Connection.Errors(0).Description & " (" & Connection.Errors(0).Source & ")"
    Else
      CCProcessError = Connection.Errors.ToString
    End If
  ElseIf NOT (Err.Description = "") Then
    CCProcessError = Err.Description
  Else
    CCProcessError = Empty
  End If
end Function
'End CCProcessError

'CCGetRequestParam @0-1DD6A561
Function CCGetRequestParam(ParameterName, Method)
  Dim ParameterValue

  If Method = ccsGet Then
    ParameterValue = Request.QueryString(ParameterName)
  ElseIf Method = ccsPost Then
    If IsMutipartEncoding Then
      ParameterValue = objUpload.Form(ParameterName)
      If Len(ParameterValue) = 0 Then 
        If Not IsEmpty(objUpload.Files(ParameterName)) Then ParameterValue = objUpload.Files(ParameterName).FileName
      End If
    Else
      ParameterValue = Request.Form(ParameterName)
    End If
  End If
  If CStr(ParameterValue) = "" Then _
    ParameterValue = Empty

  CCGetRequestParam = ParameterValue
End Function


Function CCGetRequestMultipleParam(ParameterName, Method)
  Dim ParameterValues(), ParamCount, i 

  If Method = ccsGet Then
    ParamCount = Request.QueryString(ParameterName).Count
    ReDim ParameterValues (ParamCount)
    For i = 1 To ParamCount
      ParameterValues(i) = Request.QueryString(ParameterName)(i)
      If CStr(ParameterValues(i)) = "" Then ParameterValues(i) = Empty
    Next
  ElseIf Method = ccsPost Then

    If IsMutipartEncoding Then
      Dim TempArray
      TempArray = Split(objUpload.Form(ParameterName), ", ")
      ParamCount = UBound(TempArray) + 1
      ReDim ParameterValues (ParamCount)
      For i = 0 to ParamCount - 1
        ParameterValues(i+1) = TempArray(i)
      Next
    Else
      ParamCount = Request.Form(ParameterName).Count
      ReDim ParameterValues (ParamCount)
      For i = 1 To ParamCount
        ParameterValues(i) = Request.Form(ParameterName)(i)
        If CStr(ParameterValues(i)) = "" Then ParameterValues(i) = Empty
      Next
    End If

  End If

  CCGetRequestMultipleParam = ParameterValues
End Function
'End CCGetRequestParam

'CCIsDefined @0-519FFE4F
  Function CCIsDefined(ParameterName, Scope)
   Select Case Scope
     Case "URL" 
        CCIsDefined = Not IsEmpty(Request.QueryString(ParameterName))
     Case "Form","Control"
        If IsMutipartEncoding Then
    	  CCIsDefined = Not IsEmpty(objUpload.Form(ParameterName))
	Else
    	  CCIsDefined = Not IsEmpty(Request.Form(ParameterName))
    	End If
     Case "Session"
    	CCIsDefined = Not IsEmpty(Session(ParameterName))
     Case "Application"
        CCIsDefined = Not IsEmpty(Application(ParameterName))
     Case "Cookie"
        CCIsDefined = Request.Cookies(ParameterName).HasKeys
     Case Else
        CCIsDefined = True
     End Select
  End Function
'End CCIsDefined

'CCGetQueryString @0-CC468D0D
Function CCGetQueryString(CollectionName, RemoveParameters)
  Dim QueryString, PostData, DuplicatedElements
  
  If CollectionName = "Form" Then
    QueryString = CCCollectionToString(Request.Form, RemoveParameters)
  ElseIf CollectionName = "QueryString" Then
    QueryString = CCCollectionToString(Request.QueryString, RemoveParameters)
  ElseIf CollectionName = "All" Then
    Dim RemoveParametersArray
    If TypeName(RemoveParameters) = "Variant()" Then RemoveParametersArray = RemoveParameters _
    Else RemoveParametersArray = Split(RemoveParameters, ";")
    QueryString = CCCollectionToString(Request.QueryString, RemoveParametersArray)
    DuplicatedElements = CCGetDuplicatedElementsNames(Request.Form, Request.QueryString)
    PostData = CCCollectionToString(Request.Form, IIf(Join(RemoveParametersArray, ";")<>"" And DuplicatedElements<>"", Split(Join(RemoveParametersArray, ";")+";"+DuplicatedElements, ";"), RemoveParametersArray))
    If Len(PostData) > 0 and Len(QueryString) > 0 Then _
      QueryString = QueryString & "&" & PostData _
    Else _
      QueryString = QueryString & PostData
  Else
    Err.Raise 1050, "Common Functions. CCGetQueryString Function", _
      "The CollectionName contains an illegal value."
  End If

  CCGetQueryString = QueryString
End Function
'End CCGetQueryString

'CCDuplicateElementsNames @0-930C346B
Function CCGetDuplicatedElementsNames(ParametersCollection1, ParametersCollection2)
  Dim ItemName, ItemValue, Result, Remove, I

  For Each ItemName In ParametersCollection1
    If ParametersCollection2(ItemName).Count > 0 Then
      Result = Result & ";" & ItemName
    End If
  Next

  If Len(Result) > 0 Then _
    Result = Mid(Result, 2)
  CCGetDuplicatedElementsNames = Result
End Function
'End CCDuplicateElementsNames

'CCCollectionToString @0-92DD4E55
Function CCCollectionToString(ParametersCollection, RemoveParameters)
  Dim ItemName, ItemValue, Result, Remove, I

  For Each ItemName In ParametersCollection
    Remove = false
    If IsArray(RemoveParameters) Then
      For I = 0 To UBound(RemoveParameters)
        If RemoveParameters(I) = ItemName Then 
          Remove = True
          Exit For
        End If
      Next
    End If
    If Not Remove Then
      If ParametersCollection(ItemName).Count = 1 Then
        Result = Result & _
          "&" & ItemName & "=" & Server.URLEncode(ParametersCollection(ItemName))
      Else
        For Each ItemValue In ParametersCollection(ItemName)
          Result = Result & _
            "&" & ItemName & "=" & Server.URLEncode(ItemValue)
        Next
      End If
    End If
  Next

  If Len(Result) > 0 Then _
    Result = Mid(Result, 2)
  CCCollectionToString = Result
End Function
'End CCCollectionToString

'CCAddZero @0-B5648418
Function CCAddZero(Value, ResultLength)
  Dim CountZero, I

  CountZero = ResultLength - Len(Value)
  For I = 1 To CountZero
    Value = "0" & Value
  Next 
  CCAddZero = Value
End Function
'End CCAddZero

'CCGetAMPM @0-CB6EA5BF
Function CCGetAMPM(HoursNumber, AnteMeridiem, PostMeridiem)
  If HoursNumber >= 0 And HoursNumber < 12 Then
    CCGetAMPM = AnteMeridiem
  Else
    CCGetAMPM = PostMeridiem
  End If
End Function
'End CCGetAMPM

'CC12Hour @0-12B00AFF
Function CC12Hour(HoursNumber)
  If HoursNumber = 0 Then
    HoursNumber = 12
  ElseIf HoursNumber > 12 Then
    HoursNumber = HoursNumber - 12
  End If
  CC12Hour = HoursNumber 
End Function
'End CC12Hour

'CCDBFormatByType @0-531721B5
Function CCDBFormatByType(Variable)
  Dim Result
  If VarType(Variable) = vbString Then
    If LCase(Variable) = "null" Then
      Result = Variable
    Else
      Result = "'" & Variable & "'"
    End If
  Else
    Result = CStr(Variable)
  End If
  CCDBFormatByType = Result
End Function

'End CCDBFormatByType

'CCFormatDate @0-2C65B861
Function CCFormatDate(DateToFormat, FormatMask)
  Dim ResultArray(), I, Result
  If VarType(DateToFormat) = vbEmpty Then
    Result = Empty
  ElseIf VarType(DateToFormat) <> vbDate Then
    Err.Raise 4000, "Common Functions. CCFormatDate function","Type mismatch."
  ElseIf IsEmpty(FormatMask) Then
    Result = CStr(DateToFormat)
  Else
    If CCSLocales.Locale.OverrideDateFormats Then
      Select Case FormatMask(0)
        Case "LongDate" FormatMask = CCSLocales.Locale.LongDate
        Case "LongTime" FormatMask = CCSLocales.Locale.LongTime
        Case "ShortDate" FormatMask = CCSLocales.Locale.ShortDate
        Case "ShortTime" FormatMask = CCSLocales.Locale.ShortTime
        Case "GeneralDate" FormatMask = CCSLocales.Locale.GeneralDate
	Case "ReportDate" FormatMask=Split(Join(CCSLocales.Locale.ShortDate,"|") & "| |" & Join(CCSLocales.Locale.ShortTime,"|"), "|")
      End Select
    End If
    ReDim ResultArray(UBound(FormatMask))
    For I = 0 To UBound(FormatMask)
      Select Case FormatMask(I)
        Case "d" ResultArray(I) = Day(DateToFormat)
        Case "w" ResultArray(I) = Weekday(DateToFormat)
        Case "m" ResultArray(I) = Month(DateToFormat)
        Case "q" ResultArray(I) = Fix((Month(DateToFormat) + 2) / 3)
        Case "y" ResultArray(I) = (DateDiff("d", DateSerial(Year(DateToFormat), 1, 1), DateSerial(Year(DateToFormat), Month(DateToFormat), Day(DateToFormat))) + 1)
        Case "h" ResultArray(I) = CC12Hour(Hour(DateToFormat))
        Case "H" ResultArray(I) = Hour(DateToFormat)
        Case "n" ResultArray(I) = Minute(DateToFormat)
        Case "s" ResultArray(I) = Second(DateToFormat)
        Case "wi" ResultArray(I) = CCSLocales.Locale.WeekdayNarrowNames(Weekday(DateToFormat) - 1)
        Case "dd" ResultArray(I) = CCAddZero(Day(DateToFormat), 2)
        Case "ww" ResultArray(I) = (DateDiff("ww", DateSerial(Year(DateToFormat), 1, 1), DateSerial(Year(DateToFormat), Month(DateToFormat),Day(DateToFormat))) + 1)
        Case "mm" ResultArray(I) = CCAddZero(Month(DateToFormat), 2)
        Case "yy" ResultArray(I) = Right(Year(DateToFormat), 2)
        Case "hh" ResultArray(I) = CCAddZero(CC12Hour(Hour(DateToFormat)), 2)
        Case "HH" ResultArray(I) = CCAddZero(Hour(DateToFormat), 2)
        Case "nn" ResultArray(I) = CCAddZero(Minute(DateToFormat), 2)
        Case "ss" ResultArray(I) = CCAddZero(Second(DateToFormat), 2)
        Case "S" ResultArray(I) = "000"
        Case "ddd" ResultArray(I) = CCSLocales.Locale.WeekdayShortNames(Weekday(DateToFormat) - 1)
        Case "mmm" ResultArray(I) = CCSLocales.Locale.MonthShortNames(Month(DateToFormat) - 1)
        Case "A/P" ResultArray(I) = CCGetAMPM(Hour(DateToFormat), "A", "P")
        Case "a/p" ResultArray(I) = CCGetAMPM(Hour(DateToFormat), "a", "p")
        Case "dddd" ResultArray(I) = CCSLocales.Locale.WeekdayNames(Weekday(DateToFormat) - 1)
        Case "mmmm" ResultArray(I) = CCSLocales.Locale.MonthNames(Month(DateToFormat) - 1)
        Case "yyyy" ResultArray(I) = Year(DateToFormat)
        Case "AM/PM" ResultArray(I) = CCGetAMPM(Hour(DateToFormat), "AM", "PM")
        Case "am/pm" ResultArray(I) = CCGetAMPM(Hour(DateToFormat), "am", "pm")
        Case "LongDate" ResultArray(I) = FormatDateTime(DateToFormat, vbLongDate)
        Case "LongTime" ResultArray(I) = FormatDateTime(DateToFormat, vbLongTime)
        Case "ShortDate" ResultArray(I) = FormatDateTime(DateToFormat, vbShortDate)
        Case "ShortTime" ResultArray(I) = FormatDateTime(DateToFormat, vbShortTime)
        Case "GeneralDate" ResultArray(I) = FormatDateTime(DateToFormat, vbGeneralDate)
        Case "ReportDate" ResultArray(I) = FormatDateTime(DateToFormat, vbShortDate) & " " & FormatDateTime(DateToFormat, vbShortTime)      
        Case "tt" ResultArray(I) = CCGetAMPM(Hour(DateToFormat), CCSLocales.Locale.AMDesignator, CCSLocales.Locale.PMDesignator) 
        Case Else
          If Left(FormatMask(I), 1) = "\" Then _
            ResultArray(I) = Mid(FormatMask(I), 1) _
          Else
            ResultArray(I) = FormatMask(I)
      End Select
    Next
    Result = Join(ResultArray, "")
  End If
  CCFormatDate = Result
End Function
'End CCFormatDate

'CCFormatBoolean @0-635596FD
Function CCFormatBoolean(BooleanValue, arrFormat)
  Dim Result, TrueValue, FalseValue, EmptyValue

  If IsEmpty(arrFormat) Then
    Result = CStr(BooleanValue)
  Else
    TrueValue = arrFormat(0)
    FalseValue = arrFormat(1)
    EmptyValue = arrFormat(2)
    If IsEmpty(BooleanValue) Then
      Result = EmptyValue
    Else
      If BooleanValue Then _
        Result = TrueValue _
      Else _
        Result = FalseValue
    End If
  End If
  CCFormatBoolean = Result
End Function
'End CCFormatBoolean

'CCFormatNumber @0-C9E4493E
Function CCFormatNumber(NumberToFormat, FormatArray)
  Dim IsNegative
  Dim IsExtendedFormat, IsDecimalSeparator, DecimalSeparator, IsPeriodSeparator, PeriodSeparator
  Dim DefaultDecimal, LeftPart, RightPart

  If IsEmpty(NumberToFormat) Then
    CCFormatNumber = ""
    Exit Function
  End If

  If IsArray(FormatArray) Then
    IsExtendedFormat = FormatArray(0)
    IsNegative = (NumberToFormat < 0)
    NumberToFormat = ABS(NumberToFormat) * FormatArray(7)
  
    If IsExtendedFormat Then ' Extended format
      IsDecimalSeparator = FormatArray(1)
      IsPeriodSeparator = FormatArray(3)  

      If CCSLocales.Locale.OverrideNumberFormats Then 
        DecimalSeparator = CCSLocales.Locale.DecimalSeparator
        PeriodSeparator = CCSLocales.Locale.GroupSeparator
      Else 
        DecimalSeparator = FormatArray(2)
        PeriodSeparator = FormatArray(4)
      End If

      Dim BeforeDecimal, AfterDecimal
      Dim ObligatoryBeforeDecimal, DigitsBeforeDecimal, ObligatoryAfterDecimal, DigitsAfterDecimal
      Dim I, Z
      BeforeDecimal = FormatArray(5)
      AfterDecimal = FormatArray(6)
      If IsArray(BeforeDecimal) Then
        For I = 0 To UBound(BeforeDecimal)
          If BeforeDecimal(I) = "0" Then
            ObligatoryBeforeDecimal = ObligatoryBeforeDecimal + 1
            DigitsBeforeDecimal = DigitsBeforeDecimal + 1
          ElseIf BeforeDecimal(I) = "#" Then
            DigitsBeforeDecimal = DigitsBeforeDecimal + 1
          End If
        Next      
      End If 

      If CCSLocales.Locale.OverrideNumberFormats And IsArray(AfterDecimal) Then 
        ReDim Preserve AfterDecimal(CCSLocales.Locale.DecimalDigits)
        For I = 0 To UBound(AfterDecimal)
          If AfterDecimal(I) = "" Then _
            AfterDecimal(I)="0"
        Next
      End If 

      If IsArray(AfterDecimal) Then
        For I = 0 To UBound(AfterDecimal)
          If AfterDecimal(I) = "0" Then
            ObligatoryAfterDecimal = ObligatoryAfterDecimal + 1
            DigitsAfterDecimal = DigitsAfterDecimal + 1
          ElseIf AfterDecimal(I) = "#" Then
            DigitsAfterDecimal = DigitsAfterDecimal + 1
          End If
        Next      
      End If 
  
      Dim Result, DefaultValue

      NumberToFormat = FormatNumber(NumberToFormat, DigitsAfterDecimal, False, False, False)

      DefaultDecimal = Mid(FormatNumber(10001/10, 1, True, False, True), 6, 1)
      If Not InStr(CStr(NumberToFormat), DefaultDecimal) = 0 Then
        Dim NumberParts : NumberParts = Split(CStr(NumberToFormat), DefaultDecimal)
        LeftPart = CStr(NumberParts(0))
        RightPart = CStr(NumberParts(1))
      Else
        LeftPart = CStr(NumberToFormat)
      End If

      Dim J : J = Len(LeftPart)
    
      If IsDecimalSeparator And DecimalSeparator = "" Then
        DefaultValue = CStr(FormatNumber(10001/10, 1, True, False, True))
        DecimalSeparator = Mid(DefaultValue, 6, 1)
      End If
    
      If IsPeriodSeparator And PeriodSeparator = "" Then
        DefaultValue = CStr(FormatNumber(10001/10, 1, True, False, True))
        PeriodSeparator = Mid(DefaultValue, 2, 1)
      End If  
    
      If IsArray(BeforeDecimal) Then
        Dim RankNumber : RankNumber = 0
        For I  = UBound(BeforeDecimal) To 0 Step -1
          If BeforeDecimal(i) = "#" Or BeforeDecimal(i) = "0" Then
            If DigitsBeforeDecimal = 1 And J > 1 Then
              If Not IsPeriodSeparator Then
                Result = Left(LeftPart, j) & Result
              Else
                For z = J To 1 Step -1
                  RankNumber = RankNumber + 1
                  If RankNumber Mod 3 = 1 And RankNumber - 3 > 0 Then
                    Result = Mid(LeftPart, z, 1) & PeriodSeparator & Result
                  Else
                    Result = Mid(LeftPart, z, 1) & Result
                  End If
                Next
              End If
            ElseIf J > 0 Then
              RankNumber = RankNumber + 1
              If RankNumber Mod 3 = 1 And RankNumber - 3 > 0 And IsPeriodSeparator Then
                Result = Mid(LeftPart, j, 1) & PeriodSeparator & Result
              Else
                Result = Mid(LeftPart, j, 1) & Result
              End If
              J = J - 1
              ObligatoryBeforeDecimal = ObligatoryBeforeDecimal - 1
              DigitsBeforeDecimal = DigitsBeforeDecimal - 1
            Else
              If ObligatoryBeforeDecimal > 0 Then
                RankNumber = RankNumber + 1
                If RankNumber Mod 3 = 1 And RankNumber - 3 > 0 And IsPeriodSeparator Then
                  Result = "0" & PeriodSeparator & Result
                Else
                  Result = "0" & Result
                End If
                ObligatoryBeforeDecimal = ObligatoryBeforeDecimal - 1
                DigitsBeforeDecimal = DigitsBeforeDecimal - 1
              End If
            End If
          Else
            BeforeDecimal(I) = Replace(BeforeDecimal(I), "##", "#")
            BeforeDecimal(I) = Replace(BeforeDecimal(I), "00", "0")
            Result = BeforeDecimal(I) & Result
          End If
        Next
      End If
    
      ' Left part after decimal
      Dim RightResult, IsRightResult : RightResult = "" : IsRightResult = False
      If IsArray(AfterDecimal) Then
        Dim IsZero : IsZero = True
        For I = UBound(AfterDecimal) To 0 Step -1
          If AfterDecimal(I) = "#" Or AfterDecimal(I) = "0" Then
            If DigitsAfterDecimal > ObligatoryAfterDecimal Then
              If Not Mid(RightPart, DigitsAfterDecimal, 1) = "0" Then IsZero = False
              If Not IsZero Then 
                RightResult = Mid(RightPart, DigitsAfterDecimal, 1) & RightResult
                IsRightResult = True
              End If
              DigitsAfterDecimal = DigitsAfterDecimal - 1
            Else
              RightResult = Mid(RightPart, DigitsAfterDecimal, 1) & RightResult
              DigitsAfterDecimal = DigitsAfterDecimal - 1
              IsRightResult = True
            End If
          Else
            AfterDecimal(I) = Replace(AfterDecimal(I), "##", "#")
            AfterDecimal(I) = Replace(AfterDecimal(I), "00", "0")
            RightResult = AfterDecimal(I) & RightResult
          End If
        Next
      End If

      If IsRightResult Then Result = Result & DecimalSeparator
      Result = Result & RightResult

      If NOT FormatArray(10) AND IsNegative Then _
         Result = "-" & Result

    Else ' Simple format

      If CCSLocales.Locale.OverrideNumberFormats And CInt(FormatArray(1)) <> 0 Then 
        FormatArray(1) = CCSLocales.Locale.DecimalDigits
      End If 

      If Not FormatArray(3) AND IsNegative Then _
        Result = "-" & FormatArray(5) & FormatNumber(NumberToFormat, FormatArray(1), FormatArray(2), False, FormatArray(4)) & FormatArray(6) _
      Else _
        Result = FormatArray(5) & FormatNumber(NumberToFormat, FormatArray(1), FormatArray(2), False, FormatArray(4)) & FormatArray(6)


      If CCSLocales.Locale.OverrideNumberFormats Then 
        DefaultDecimal = Mid(FormatNumber(10001/10, 1, True, False, True), 6, 1)
        If InStr(CStr(Result), DefaultDecimal) > 0 Then
          Result = Split(CStr(Result), DefaultDecimal)
        End If
        If FormatArray(4) Then 
           DefaultValue = CStr(FormatNumber(10001/10, 1, True, False, True))
           PeriodSeparator = Mid(DefaultValue, 2, 1)
           If IsArray(Result) Then 
             Result(0) = Replace(Result(0), PeriodSeparator, CCSLocales.Locale.GroupSeparator) 
           Else 
             Result = Replace(Result, PeriodSeparator, CCSLocales.Locale.GroupSeparator) 
           End If
        End If
        If IsArray(Result) Then _
          Result = Join(Result, CCSLocales.Locale.DecimalSeparator)
      End If
    End If
    If Not FormatArray(8) Then Result = Server.HTMLEncode(Result)
    If Not CStr(FormatArray(9)) = "" Then _
      Result = "<FONT COLOR=""" & FormatArray(9) & """>" & Result & "</FONT>"
  Else
    Result = CStr(NumberToFormat)
  End If
  CCFormatNumber = Result

End Function
'End CCFormatNumber

'CCParseBoolean @0-33711A62
Function CCParseBoolean(Value, FormatMask)
  Dim Result
  Result = Empty
  If VarType(Value) = vbBoolean Then
    Result = Value
  Else
    If IsEmpty(FormatMask) Then
      Result = CBool(Value)
    Else
      If IsEmpty(Value) Then
        If CStr(FormatMask(0)) = "null" Then _
          Result = True
        If CStr(FormatMask(1)) = "null" Then _
          Result = False
      Else
        If CStr(Value) = CStr(FormatMask(0)) Then 
          Result = True
        ElseIf CStr(Value) = CStr(FormatMask(1)) Then
          Result = False
        End If
      End If
    End If
  End If
  CCParseBoolean = Result
End Function
'End CCParseBoolean

'CCParseDate @0-B0351854
Function CCParseDate(ParsingDate, FormatMask)
  Dim ResultDate, ResultDateArray(8)
  Dim MaskPart, MaskLength, TokenLength
  Dim IsError
  Dim DatePosition, MaskPosition
  Dim Delimiter, BeginDelimiter
  Dim MonthNumber, MonthName, MonthArray
  Dim DatePartStr

  Dim IS_DATE_POS, YEAR_POS, MONTH_POS, DAY_POS, IS_TIME_POS, HOUR_POS, MINUTE_POS, SECOND_POS

  IS_DATE_POS = 0 : YEAR_POS = 1 : MONTH_POS = 2 : DAY_POS = 3
  IS_TIME_POS = 4 : HOUR_POS = 5 : MINUTE_POS = 6 : SECOND_POS = 7

  If VarType(ParsingDate) = vbDate Then 
     CCParseDate = ParsingDate
     Exit Function
  End If


  If IsEmpty(FormatMask) Then
    If CStr(ParsingDate) = "" Then _
      ResultDate = Empty _
    Else _
      ResultDate = CDate(ParsingDate)
  ElseIf CStr(ParsingDate) = "" Then
    ResultDate = Empty
  Else
    If CCSLocales.Locale.OverrideDateFormats Then
      Select Case FormatMask(0)
        Case "LongDate" FormatMask = CCSLocales.Locale.LongDate
        Case "LongTime" FormatMask = CCSLocales.Locale.LongTime
        Case "ShortDate" FormatMask = CCSLocales.Locale.ShortDate
        Case "ShortTime" FormatMask = CCSLocales.Locale.ShortTime
        Case "GeneralDate" FormatMask = CCSLocales.Locale.GeneralDate
      End Select
    ElseIf (FormatMask(0) = "GeneralDate" Or FormatMask(0) = "LongDate" _
      Or FormatMask(0) = "ShortDate" Or FormatMask(0) = "LongTime" _ 
      Or FormatMask(0) = "ShortTime") And Not CStr(ParsingDate) = "" Then
         If Not IsDate(ParsingDate) Then  Err.Raise 4000, "Common Functions. ParseDate function", "Mask mismatch."  
         CCParseDate = CDate(ParsingDate)
         Exit Function
    End If
    DatePosition = 1
    MaskPosition = 0
    MaskLength = UBound(FormatMask)
    IsError = False

    ' Default date
    ResultDateArray(IS_DATE_POS) = False
    ResultDateArray(IS_TIME_POS) = False
    ResultDateArray(YEAR_POS) = 0 : ResultDateArray(MONTH_POS) = 12 : ResultDateArray(DAY_POS) = 1
    ResultDateArray(HOUR_POS) = 0 : ResultDateArray(MINUTE_POS) = 0 : ResultDateArray(SECOND_POS) = 0

    While (MaskPosition <= MaskLength) AND NOT IsError
      MaskPart = FormatMask(MaskPosition)
      If CCSDateConstants.DateMasks.Exists(MaskPart) Then
        TokenLength = CCSDateConstants.DateMasks(MaskPart)
        If TokenLength > 0 Then
          DatePartStr = Mid(ParsingDate, DatePosition, TokenLength)
          DatePosition = DatePosition + TokenLength
        Else
          If MaskPosition < MaskLength Then
            Delimiter = FormatMask(MaskPosition + 1)
            BeginDelimiter = InStr(DatePosition, ParsingDate, Delimiter)
            If BeginDelimiter = 0 Then
              Err.Raise 4000, "Common Functions. ParseDate function","Mask mismatch."
            Else
              DatePartStr = Mid(ParsingDate, DatePosition, BeginDelimiter - DatePosition)
              DatePosition = BeginDelimiter
            End If
          Else
            DatePartStr = Mid(ParsingDate, DatePosition)
            DatePosition = DatePosition &  Len(DatePartStr)
          End If
        End If
        Select Case MaskPart
          Case "d", "dd"
            ResultDateArray(DAY_POS) = CInt(DatePartStr)
            ResultDateArray(IS_DATE_POS) = True
          Case "ddd", "dddd"
            Dim DayArray, DayNumber, DayName
            DayNumber = 0
            DayName = UCase(DatePartStr)
            If MaskPart = "ddd" Then _
              DayArray = CCSLocales.Locale.WeekdayShortNames _
            Else _
              DayArray = CCSLocales.Locale.WeekdayNames
            While DayNumber < 6 AND UCase(DayArray(DayNumber)) <> DayName
              DayNumber = DayNumber + 1
            Wend
            If DayNumber = 6 Then
            If UCase(DayArray(6)) <> DayName Then _
              Err.Raise 4000, "Common Functions. ParseDate function","Mask mismatch."
            End If
          Case "m", "mm"
            ResultDateArray(MONTH_POS) = CInt(DatePartStr)
            ResultDateArray(IS_DATE_POS) = True
          Case "mmm", "mmmm"
            MonthNumber = 0
            MonthName = UCase(DatePartStr)
            If MaskPart = "mmm" Then _
              MonthArray = CCSLocales.Locale.MonthShortNames _
            Else _
              MonthArray = CCSLocales.Locale.MonthNames
            While MonthNumber < 11 AND UCase(MonthArray(MonthNumber)) <> MonthName
              MonthNumber = MonthNumber + 1
            Wend
            If MonthNumber = 11 Then
              If UCase(MonthArray(11)) <> MonthName Then _
                Err.Raise 4000, "Common Functions. ParseDate function", "Mask mismatch."
            End If
            ResultDateArray(MONTH_POS) = MonthNumber + 1
            ResultDateArray(IS_DATE_POS) = True
          Case "yyyy"
            ResultDateArray(YEAR_POS) = CInt(DatePartStr)
            ResultDateArray(IS_DATE_POS) = True
          Case "yy"
            If CInt(DatePartStr) >= 50 Then ResultDateArray(YEAR_POS) = 1900 + CInt(DatePartStr) _
            Else ResultDateArray(YEAR_POS) = 2000 + CInt(DatePartStr)
            ResultDateArray(IS_DATE_POS) = True
          Case "h", "hh"
            If CInt(DatePartStr) = 12 Then _
              ResultDateArray(HOUR_POS) = 0 _
            Else _
              ResultDateArray(HOUR_POS) = CInt(DatePartStr)
            ResultDateArray(IS_TIME_POS) = True
          Case "H", "HH"
            ResultDateArray(HOUR_POS) = CInt(DatePartStr)
            ResultDateArray(IS_TIME_POS) = True
          Case "n", "nn"
            ResultDateArray(MINUTE_POS) = CInt(DatePartStr)
            ResultDateArray(IS_TIME_POS) = True
          Case "s", "ss"
            ResultDateArray(SECOND_POS) = CInt(DatePartStr)
            ResultDateArray(IS_TIME_POS) = True
          Case "am/pm", "a/p", "AM/PM", "A/P"
            If Left(LCase(DatePartStr), 1) = "p" Then
              ResultDateArray(HOUR_POS) = ResultDateArray(HOUR_POS) + 12
            ElseIf Left(LCase(DatePartStr), 1) = "a" Then
              ResultDateArray(HOUR_POS) = ResultDateArray(HOUR_POS)
            End If
            ResultDateArray(IS_TIME_POS) = True
          Case "tt" 
            If DatePartStr = CCSLocales.Locale.PMDesignator Then _ 
              ResultDateArray(HOUR_POS) = ResultDateArray(HOUR_POS) + 12
            ResultDateArray(IS_TIME_POS) = True
          Case "w", "q","S"
            ' Do Nothing
          Case Else
            IsError = IsError And DatePartStr = MaskPart
        End Select
      Else
        DatePartStr = Mid(ParsingDate, DatePosition, Len(FormatMask(MaskPosition)))
        DatePosition = DatePosition + Len(FormatMask(MaskPosition))
        If FormatMask(MaskPosition) <> DatePartStr Then _
          IsError = True
      End If
      MaskPosition = MaskPosition + 1
    Wend

    If Len(ParsingDate) - DatePosition >= 0  Then IsError = True
    If IsError Then Err.Raise 4001, "Common Functions. CCParseDate Function", "Unable to parse the date value."

    If ResultDateArray(IS_DATE_POS) AND ResultDateArray(IS_TIME_POS) Then
      ResultDate = DateSerial(ResultDateArray(YEAR_POS), ResultDateArray(MONTH_POS), ResultDateArray(DAY_POS))
      ResultDate = DateAdd("h", ResultDateArray(HOUR_POS), ResultDate)
      ResultDate = DateAdd("n", ResultDateArray(MINUTE_POS), ResultDate)
      ResultDate = DateAdd("s", ResultDateArray(SECOND_POS), ResultDate)
      If NOT(Year(ResultDate) = ResultDateArray(YEAR_POS) _
        AND Month(ResultDate) = ResultDateArray(MONTH_POS) _
        AND Day(ResultDate) = ResultDateArray(DAY_POS) _
        AND Hour(ResultDate) = ResultDateArray(HOUR_POS) _
        AND Minute(ResultDate) = ResultDateArray(MINUTE_POS) _
        AND Second(ResultDate) = ResultDateArray(SECOND_POS)) _
      Then _
        Err.Raise 4001,"Common Functions. CCParseDate Function", "Unable to parse the date value."
    ElseIf ResultDateArray(IS_TIME_POS) Then 
      ResultDate = TimeSerial(ResultDateArray(HOUR_POS), ResultDateArray(MINUTE_POS), ResultDateArray(SECOND_POS))
      If NOT(Hour(ResultDate) = ResultDateArray(HOUR_POS) _
        AND Minute(ResultDate) = ResultDateArray(MINUTE_POS) _
        AND Second(ResultDate) = ResultDateArray(SECOND_POS)) _
      Then _
        Err.Raise 4001,"Common Functions. CCParseDate Function", "Unable to parse the date value."
    ElseIf ResultDateArray(IS_DATE_POS) Then
      ResultDate = DateSerial(ResultDateArray(YEAR_POS), ResultDateArray(MONTH_POS), ResultDateArray(DAY_POS))
      If NOT(Year(ResultDate) = ResultDateArray(YEAR_POS) _
        AND Month(ResultDate) = ResultDateArray(MONTH_POS) _
        AND Day(ResultDate) = ResultDateArray(DAY_POS)) _
      Then _
        Err.Raise 4001, "Common Functions. CCParseDate Function", "Unable to parse the date value."
    End If
  End If
  CCParseDate = ResultDate
End Function
'End CCParseDate

'CCParseNumber @0-B8E8F682
Function CCParseNumber(NumberValue, FormatArray, DataType)
  Dim Result, NumberValueType, NumberVal
  NumberValueType = VarType(NumberValue)
  If NumberValueType = vbInteger OR NumberValueType = vbLong _
    OR NumberValueType = vbSingle OR NumberValueType = vbSingle _
    OR NumberValueType = vbCurrency OR NumberValueType = vbDecimal _
    OR NumberValueType = vbByte Then
    If DataType = ccsInteger Then
      Result = CLng(NumberValue)
    ElseIf DataType = ccsFloat Then
      Result = CDbl(NumberValue)
    ElseIf DataType = ccsSingle Then
      Result = CSng(NumberValue)
    End If
  Else
    If Not CStr(NumberValue) = "" Then
      Dim DefaultValue, DefaultDecimal
      Dim DecimalSeparator, PeriodSeparator, PrePart, PostPart
      DecimalSeparator = "" : PeriodSeparator = "" : PrePart="" : PostPart=""
      If IsArray(FormatArray) Then
        If FormatArray(0) Then
          If CCSLocales.Locale.OverrideNumberFormats Then 
            DecimalSeparator = CCSLocales.Locale.DecimalSeparator
            PeriodSeparator = CCSLocales.Locale.GroupSeparator
          Else 
            DecimalSeparator = FormatArray(2)
            PeriodSeparator = FormatArray(4)
          End If
        Else
          PrePart = FormatArray(5)
          PostPart = FormatArray(6)
        End If
      End If
      NumberVal = NumberValue
      If Not CStr(DecimalSeparator) = "" Then 
        DefaultValue = CStr(FormatNumber(10001/10, 1, True, False, True))
        DefaultDecimal = Mid(DefaultValue, 6, 1)
        NumberVal = Replace(NumberVal, DecimalSeparator, DefaultDecimal)
      End If
      If Not CStr(PeriodSeparator) = "" Then NumberVal = Replace(NumberVal, PeriodSeparator, "")
      If Not CStr(PrePart) = "" Then NumberVal = Replace(NumberVal, PrePart, "")
      If Not CStr(PostPart) = "" Then NumberVal = Replace(NumberVal, PostPart, "")
      If DataType = ccsInteger Then
        Result = CLng(NumberVal)
      ElseIf DataType = ccsFloat Then
        Result = CDbl(NumberVal)
      ElseIf DataType = ccsSingle Then
        Result = CSng(NumberVal)
      End If
      If IsArray(FormatArray) Then Result = Result/FormatArray(7)
    Else
      Result = Empty
    End If
  End If
  CCParseNumber = Result
End Function
'End CCParseNumber

'CCParseInteger @0-42815927
Function CCParseInteger(NumberValue, FormatArray)
  CCParseInteger = CCParseNumber(NumberValue, FormatArray, ccsInteger)
End Function
'End CCParseInteger

'CCParseFloat @0-56667DF0
Function CCParseFloat(NumberValue, FormatArray)
  CCParseFloat = CCParseNumber(NumberValue, FormatArray, ccsFloat)
End Function
'End CCParseFloat

'CCParseSingle @0-0142EA0D
Function CCParseSingle(NumberValue, FormatArray)
  CCParseSingle = CCParseNumber(NumberValue, FormatArray, ccsSingle)
End Function
'End CCParseSingle

'CCValidateDate @0-B1691F92
Function CCValidateDate(ValidatingDate, FormatMask)
  Dim MaskPosition, I, Result, OneChar, IsSeparator
  Dim RegExpPattern, RegExpObject, Matches
  Dim ParsedTestDate, FormattedTestDate

  IsSeparator = False

  If ValidatingDate = "" OR IsEmpty(ValidatingDate) Then
    Result = True
  ElseIf IsEmpty(FormatMask) Then
    Result = IsDate(ValidatingDate)
  Else
    If CCSLocales.Locale.OverrideDateFormats Then
      Select Case FormatMask(0)
        Case "LongDate" FormatMask = CCSLocales.Locale.LongDate
        Case "LongTime" FormatMask = CCSLocales.Locale.LongTime
        Case "ShortDate" FormatMask = CCSLocales.Locale.ShortDate
        Case "ShortTime" FormatMask = CCSLocales.Locale.ShortTime
        Case "GeneralDate" FormatMask = CCSLocales.Locale.GeneralDate
      End Select
    ElseIf FormatMask(0) = "GeneralDate" Or FormatMask(0) = "LongDate" _
       Or FormatMask(0) = "ShortDate" Or FormatMask(0) = "LongTime" _ 
       Or FormatMask(0) = "ShortTime" Then
       CCValidateDate = IsDate(ValidatingDate)
       Exit Function
    End If
    ParsedTestDate = CCParseDate(ValidatingDate, FormatMask)
    FormattedTestDate = CCFormatDate(ParsedTestDate, FormatMask)
    Result = FormattedTestDate = ValidatingDate
  End If
  CCValidateDate = Result
End Function
'End CCValidateDate

'CCValidateNumber @0-08089509
Function CCValidateNumber(NumberValue, FormatArray)
  Dim Result, NumberValueType
  NumberValueType = VarType(NumberValue)
  If NumberValueType = vbInteger OR NumberValueType = vbLong _
    OR NumberValueType = vbSingle OR NumberValueType = vbSingle _
    OR NumberValueType = vbCurrency OR NumberValueType = vbDecimal _
    OR NumberValueType = vbByte Then
      Result = True
  Else
    If Not CStr(NumberValue) = "" Then
      Dim DefaultValue, DefaultDecimal
      Dim DecimalSeparator, PeriodSeparator
      DecimalSeparator = "" : PeriodSeparator = ""
      If IsArray(FormatArray) Then
        If FormatArray(0) Then
          DecimalSeparator = FormatArray(2)
          PeriodSeparator = FormatArray(4)
        End If
      End If
      If Not CStr(DecimalSeparator) = "" Then 
        DefaultValue = CStr(FormatNumber(10001/10, 1, True, False, True))
        DefaultDecimal = Mid(DefaultValue, 6, 1)
        NumberValue = Replace(NumberValue, DecimalSeparator, DefaultDecimal)
      End If
      If Not CStr(PeriodSeparator) = "" Then NumberValue = Replace(NumberValue, PeriodSeparator, "")
      Result = IsNumeric(NumberValue)
    Else
      Result = True
    End If
  End If
  CCValidateNumber = Result
End Function
'End CCValidateNumber

'CCValidateBoolean @0-B8DE2060
Function CCValidateBoolean(Value, FormatMask)
  Dim Result: Result = False

  If VarType(Value) = vbBoolean Then
    Result = True
  Else
    If IsEmpty(FormatMask) Then
      On Error Resume Next
      Result = CBool(Value)
      Result = Not(Err > 0)
    Else
      If IsEmpty(Value) Or CStr(Value) = "" Then
        Result = (CStr(FormatMask(0)) = "null") Or (CStr(FormatMask(0)) = "Undefined") Or (CStr(FormatMask(0)) = "")
        Result = Result Or (CStr(FormatMask(1)) = "null") Or (CStr(FormatMask(1)) = "Undefined") Or (CStr(FormatMask(1)) = "")
        If UBound(FormatMask) = 2 Then _
          Result = Result Or (CStr(FormatMask(2)) = "null") Or (CStr(FormatMask(2)) = "Undefined") Or (CStr(FormatMask(2)) = "")
      Else
        Result = (CStr(Value) = CStr(FormatMask(0))) Or (CStr(Value) = CStr(FormatMask(1)))
        If UBound(FormatMask) = 2 Then _
          Result = Result Or (CStr(Value) = CStr(FormatMask(2)))
      End If
    End If
  End If
  CCValidateBoolean = Result
End Function
'End CCValidateBoolean

'CCAddParam @0-54F4416E
Function CCAddParam(QueryString, ParameterName, ParameterValue)
  Dim Result, ParameterValues, i, j, re

  Result = QueryString
  i = InStr(LCase(Result), LCase(ParameterName) & "=")
  If i > 1 Then i = InStr(LCase(Result), "&" & LCase(ParameterName) & "=")
  If i > 0 Then j = InStr(i + 1, LCase(Result), "&")
  While i > 0
    If j > 0 Then Result = Mid(Result, 1, i - 1) + Mid(Result, j) Else Result = Mid(Result, 1, i - 1)
    i = InStr(LCase(Result), "&" & LCase(ParameterName) & "=")
    If i > 0 Then j = InStr(i + 1, LCase(Result), "&")
  WEnd
  ParameterValues = Split(CStr(ParameterValue), ", ")
  If UBound(ParameterValues) > 0 Then
    For i = 0 To UBound(ParameterValues)
      Result = Result & "&" & ParameterName & "=" & Server.URLEncode(ParameterValues(i))
    Next
  Else
    Result = Result & "&" & ParameterName & "=" & Server.URLEncode(ParameterValue)
  End If
  Result = Replace(Result, "&&", "&")
  If Left(Result, 1) = "&" Then Result = Mid(Result, 2)
  CCAddParam = Result
End Function
'End CCAddParam

'CCRemoveParam @0-C58D1DC2
Function CCRemoveParam(QueryString, ParameterName)
  Dim Result
  Result = Replace(QueryString, ParameterName & "=" & Server.URLEncode(CCGetFromGet(ParameterName, Empty)), "", 1, -1, 1)
  Result = Replace(Result, "&&", "&")
  If Left(Result, 1) = "&" Then Result = Mid(Result, 2)
  CCRemoveParam = Result
End Function
'End CCRemoveParam

'CCRegExpTest @0-9EAA5A2D
Function CCRegExpTest(TestValue, RegExpMask, IgnoreCase, GlobalTest)
  Dim Result
  If Not CStr(TestValue) = "" Then
    Dim RegExpObject
    Set RegExpObject = New RegExp
    RegExpObject.Pattern = RegExpMask
    RegExpObject.IgnoreCase = IgnoreCase
    RegExpObject.Global = GlobalTest
    Result = RegExpObject.Test(CStr(TestValue)) 
    Set RegExpObject = Nothing
  Else
    Result = True    
  End If
  CCRegExpTest = Result
End Function
  

'End CCRegExpTest

'CCRegExpReplace @0-C56ABB12
Function CCRegExpReplace(TestValue, RegExpMask, NewValue,IgnoreCase)
  Dim Result
  If Not CStr(TestValue) = "" Then
    Dim RegExpObject
    Set RegExpObject = New RegExp
    RegExpObject.Pattern = RegExpMask
    RegExpObject.IgnoreCase = IgnoreCase
    Result = RegExpObject.Replace(CStr(TestValue),CStr(NewValue)) 
    Set RegExpObject = Nothing
  Else
    Result = ""    
  End If
  CCRegExpReplace = Result
End Function


'End CCRegExpReplace

'CheckSSL @0-4BE3AE1D
Sub CheckSSL()
  If Not UCase(Request.ServerVariables("HTTPS")) = "ON" Then
    Response.Write "SSL connection error. This page can be accessed only via secured connection."
    Response.End
  End If
End Sub

'End CheckSSL

'setInclPath @0-0B2C5DD1

Function setInclPath(o, n)
 Dim aro, arn, j, path

 If o = "" Then 
    setInclPath = n
    Exit Function
 End If 

 If Right(o, 1) = "/" Then o = Left(o, Len(o) - 1)
 If Right(n, 1) = "/" Then n = Left(n, Len(n) - 1)
 aro = Split(o, "/")
 arn = Split(n, "/")

 For j = LBound(arn) To UBound(arn)
    If Left(arn(j), 2) = ".." Then
      If Left(aro(UBound(aro)), 2) = ".." Then 
        ReDim Preserve aro(UBound(aro) + 1)
        aro(UBound(aro)) = arn(j)
      Else 
         ReDim Preserve aro(UBound(aro) - 1)
      End If
   ElseIf Left(arn(j), 1) = "." Then
   ElseIf Trim(arn(j)) <> "" Then
      ReDim Preserve aro(UBound(aro) + 1)
     aro(UBound(aro)) = arn(j)
   End If
 Next
 path = Join(aro, "/")
 If path <> "" Then path = path & "/"
 setInclPath = path
End Function

'End setInclPath

'CCGetUserLogin @0-4306ED6C
Function CCGetUserLogin()
    CCGetUserLogin = Session("UserLogin")
End Function
'End CCGetUserLogin

'CCSecurityRedirect @0-2ACFBE19
Sub CCSecurityRedirect(GroupsAccess, URL)
    Dim ErrorType
    Dim RetLink
    Dim RetLinkParams
    Dim Link
    ErrorType = CCSecurityAccessCheck(GroupsAccess)
    If NOT (ErrorType = "success") Then
        If IsEmpty(URL) Then _
            Link = ServerURL & "Login.asp" _
        Else _
            Link = URL
        RetLink = Request.ServerVariables("SCRIPT_NAME")
        RetLinkParams = CCRemoveParam(Request.ServerVariables("QUERY_STRING"), "ccsForm")
        If NOT (RetLinkParams = "") Then _
            RetLink = RetLink & "?" & RetLinkParams
        Response.Redirect(Link & "?ret_link=" & _
            Server.URLEncode(RetLink) & "&type=" & ErrorType)
    End If
End Sub
'End CCSecurityRedirect

'CCGetUserID @0-449B3B19
Function CCGetUserID()
    CCGetUserID = Session("UserID")
End Function
'End CCGetUserID

'CCSecurityAccessCheck @0-8A7701BE
Function CCSecurityAccessCheck(GroupsAccess)
    Dim ErrorType
    Dim GroupID
    ErrorType = "success"
    If IsEmpty(CCGetUserID()) Then
        ErrorType = "notLogged"
    Else
        GroupID = CCGetGroupID()
        If IsEmpty(GroupID) Then
            ErrorType = "groupIDNotSet"
        Else
            If NOT CCUserInGroups(GroupID, GroupsAccess) Then
                ErrorType = "illegalGroup"
            End If
        End If
    End If
    CCSecurityAccessCheck = ErrorType
End Function
'End CCSecurityAccessCheck

'CCGetGroupID @0-B2650479
Function CCGetGroupID()
CCGetGroupID = Session("GroupID")
End Function
'End CCGetGroupID

'CCUserInGroups @0-4332AEA7
Function CCUserInGroups(GroupID, GroupsAccess)
Dim Result
Dim GroupNumber
If NOT IsEmpty(GroupsAccess) Then
GroupNumber = CLng(GroupID)
While NOT Result AND GroupNumber > 0
Result = NOT (InStr(";" & GroupsAccess & ";", ";" & GroupNumber & ";") = 0)
GroupNumber = GroupNumber - 1
Wend
Else
Result = True
End If
CCUserInGroups = Result
End Function
'End CCUserInGroups

'CCLoginUser @0-BF9977E5
Function CCLoginUser(Login, Password)
    Dim Result
    Dim SQL
    Dim RecordSet
    Dim Connection
    Dim UserIDField
    Dim GroupIDField
    
    Set Connection = New clsDBSystem
    Connection.Open
    SQL = "SELECT [UserTableID], [UserTableUserGroupID] FROM [ActiveUserView] WHERE [UserTableLogin]='" & Replace(Login, "'", "''") & "' AND [UserTablePassword]='" & Replace(Password, "'", "''") & "'"
    Set RecordSet = Connection.Execute(SQL)
    Result = NOT RecordSet.EOF
    If Result Then
        UserIDField = CStr(RecordSet("UserTableID").Value)
        Session("UserID") = UserIDField
        Session("UserLogin") = Login
        Set GroupIDField = RecordSet("UserTableUserGroupID")
        Select Case GroupIDField.Type
            Case adBigInt, adUnsignedBigInt
                GroupIDField = CLng(GroupIDField.Value)
            Case adBoolean
                GroupIDField = CBool(GroupIDField.Value)
            Case adBSTR, adChar, adVarChar, adLongVarChar, adLongVarWChar, adVarWChar, adWChar
                GroupIDField = CStr(GroupIDField.Value)
            Case adCurrency
                GroupIDField = CCur(GroupIDField.Value)
            Case adDate, adDBDate
                GroupIDField = CDate(GroupIDField.Value)
            Case adDecimal, adDouble, adNumeric, adVarNumeric
                GroupIDField = CDbl(GroupIDField.Value)
            Case adEmpty
                GroupIDField = Empty
            Case adInteger, adSmallInt, adTinyInt, adUnsignedInt, adUnsignedSmallInt
                GroupIDField = CInt(GroupIDField.Value)
            Case adSingle
                GroupIDField = CSng(GroupIDField.Value)
            Case adUnsignedTinyInt
                GroupIDField = CByte(GroupIDField.Value)
            Case Else
                GroupIDField = GroupIDField.Value
        End Select
        Session("GroupID") = GroupIDField
    End If
    RecordSet.Close
    Set RecordSet = Nothing
    Connection.Close
    Set Connection = Nothing
    CCLoginUser = Result
End Function
'End CCLoginUser

'CCLogoutUser @0-DB93CE50
Sub CCLogoutUser()
    Session("UserID") = Empty
    Session("UserLogin") = Empty
    Session("GroupID") = Empty
End Sub
'End CCLogoutUser

'GetCCSType @0-8E845BA4
Function GetCCSType(adType)
  Dim Res : Res =ccsText
  Select Case adType
   Case adBigInt
      Res = ccsInteger
   Case adChar
      Res = ccsText
   Case adDate
      Res = ccsDate
   Case adDecimal
      Res = ccsFloat
   Case adDouble
      Res = ccsFloat
   Case adNumeric
      Res = ccsFloat
   Case adSmallInt
      Res = ccsInteger
   Case adTinyInt
      Res = ccsInteger
   Case adVarChar
      Res = ccsText
   Case adBoolean
      Res = ccsBoolean
   Case adDBTimeStamp
      Res = ccsDate
   Case adInteger
      Res = ccsInteger
   Case adWChar
      Res = ccsText
   Case adBSTR
      Res = ccsText
   Case adSingle
      Res = ccsSingle
   Case adDate
      Res = ccsDate
   Case Else
      Res = ccsText
  End Select
  GetCCSType = Res
End Function
'End GetCCSType

'CCGetFormatStr @0-4618C2F1
Function CCGetFormatStr(Format)
  Dim Result
  If IsEmpty(Format) Then
    Result = ""
  Else
    Select Case Format(0)
      Case "LongDate" Result = Join(CCSLocales.Locale.LongDate, "")
      Case "LongTime" Result = Join(CCSLocales.Locale.LongTime, "")
      Case "ShortDate" Result = Join(CCSLocales.Locale.ShortDate, "")
      Case "ShortTime" Result = Join(CCSLocales.Locale.ShortTime, "")
      Case "GeneralDate" Result = Join(CCSLocales.Locale.GeneralDate, "")
      Case Else Result = Join(Format, "")
    End Select  
  End If
  CCGetFormatStr = Result
End Function
'End CCGetFormatStr

'CCLoadStaticTranslation @0-0186EB2D
  Public Function CCLoadStaticTranslation()
    Dim Keys(97)
    Dim Vals(97)
    
    Keys(1) = "ccs_asc" : Vals(1) = "Ascending"
    Keys(2) = "ccs_bytes" : Vals(2) = "bytes"
    Keys(3) = "ccs_cancel" : Vals(3) = "Cancel"
    Keys(4) = "ccs_cannotseek" : Vals(4) = "Cannot find specified record."
    Keys(5) = "ccs_clear" : Vals(5) = "Clear"
    Keys(6) = "ccs_customlinkfield" : Vals(6) = "Detail"
    Keys(7) = "ccs_customoperationerror_missingparameters" : Vals(7) = "One or more parameters missing to perform the Update/Delete. The application is misconfigured."
    Keys(8) = "ccs_databasecommanderror" : Vals(8) = "Database command error."
    Keys(9) = "ccs_datepickernav61" : Vals(9) = "Date Picker component is not compatible with Netscape 6.1"
    Keys(10) = "ccs_delete" : Vals(10) = "Delete"
    Keys(11) = "ccs_deleteconfirmation" : Vals(11) = "Delete record?"
    Keys(12) = "ccs_desc" : Vals(12) = "Descending"
    Keys(13) = "ccs_directoryformprefix" : Vals(13) = "Directory"
    Keys(14) = "ccs_directoryformsuffix" : Vals(14) = ""
    Keys(15) = "ccs_filenotfound" : Vals(15) = "The file {0} specified in {1} was not found."
    Keys(16) = "ccs_filesfoldernotfound" : Vals(16) = "Unable to upload the file specified in {0} - upload folder doesn't exist."
    Keys(17) = "ccs_filter" : Vals(17) = "Keyword"
    Keys(18) = "ccs_first" : Vals(18) = "First"
    Keys(19) = "ccs_formatinfo" : Vals(19) = "en|en|US|Yes;No;|2|.|,|January;February;March;April;May;June;July;August;September;October;November;December|Jan;Feb;Mar;Apr;May;Jun;Jul;Aug;Sep;Oct;Nov;Dec|Sunday;Monday;Tuesday;Wednesday;Thursday;Friday;Saturday|Sun;Mon;Tue;Wed;Thu;Fri;Sat|m!/!d!/!yyyy|dddd!, !mmmm! !dd!, !yyyy|h!:!nn! !tt|h!:!nn!:!ss! !tt|0|AM|PM|windows-1252|1252|0|0|S;M;T;W;T;F;S|1033"
    Keys(20) = "ccs_gridformprefix" : Vals(20) = "List of"
    Keys(21) = "ccs_gridformsuffix" : Vals(21) = ""
    Keys(22) = "ccs_gridpagenumbererror" : Vals(22) = "Invalid page number."
    Keys(23) = "ccs_gridpagesizeerror" : Vals(23) = "(CCS06) Invalid page size."
    Keys(24) = "ccs_incorrectemailformat" : Vals(24) = "Invalid email format in field {0}."
    Keys(25) = "ccs_incorrectformat" : Vals(25) = "The value in field {0} is not valid. Use the following format: {1}."
    Keys(26) = "ccs_incorrectphoneformat" : Vals(26) = "Invalid phone number format in field {0}."
    Keys(27) = "ccs_incorrectvalue" : Vals(27) = "The value in field {0} is not valid."
    Keys(28) = "ccs_incorrectzipformat" : Vals(28) = "Invalid zip code format in field {0}."
    Keys(29) = "ccs_insert" : Vals(29) = "Add"
    Keys(30) = "ccs_insertlink" : Vals(30) = "Add New"
    Keys(31) = "ccs_insufficientpermissions" : Vals(31) = "Insufficient filesystem permissions to upload the file specified in {0}."
    Keys(32) = "ccs_languageid" : Vals(32) = "en"
    Keys(33) = "ccs_largefile" : Vals(33) = "The file size in field {0} is too large."
    Keys(34) = "ccs_last" : Vals(34) = "Last"
    Keys(35) = "ccs_localeid" : Vals(35) = "en"
    Keys(36) = "ccs_login" : Vals(36) = "Login"
    Keys(37) = "ccs_loginbtn" : Vals(37) = "Login"
    Keys(38) = "ccs_loginerror" : Vals(38) = "Login or Password is incorrect."
    Keys(39) = "ccs_login_form_caption" : Vals(39) = "Login"
    Keys(40) = "ccs_logoutbtn" : Vals(40) = "Logout"
    Keys(41) = "ccs_main" : Vals(41) = "Main"
    Keys(42) = "ccs_maskvalidation" : Vals(42) = "Mask validation failed for field {0}."
    Keys(43) = "ccs_maximumlength" : Vals(43) = "The number of symbols in field {0} can't be greater than {1}."
    Keys(44) = "ccs_maximumvalue" : Vals(44) = "The value in field {0} can't be greater than {1}."
    Keys(45) = "ccs_minimumlength" : Vals(45) = "The number of symbols in field {0} can't be less than {1}."
    Keys(46) = "ccs_minimumvalue" : Vals(46) = "The value in field {0} can't be less than {1}."
    Keys(47) = "ccs_more" : Vals(47) = "More..."
    Keys(48) = "ccs_next" : Vals(48) = "Next"
    Keys(49) = "ccs_nextmonthhint" : Vals(49) = "Next Month"
    Keys(50) = "ccs_nextquarterhint" : Vals(50) = "Next Quarter"
    Keys(51) = "ccs_nextthreemonthshint" : Vals(51) = "Next Three Months"
    Keys(52) = "ccs_nextyearhint" : Vals(52) = "Next Year"
    Keys(53) = "ccs_nocategories" : Vals(53) = "No categories found"
    Keys(54) = "ccs_norecords" : Vals(54) = "No records"
    Keys(55) = "ccs_of" : Vals(55) = "of"
    Keys(56) = "ccs_operationerror" : Vals(56) = "Unable to perform the {0} operation. One or more parameters are unspecified."
    Keys(57) = "ccs_password" : Vals(57) = "Password"
    Keys(58) = "ccs_previous" : Vals(58) = "Prev"
    Keys(59) = "ccs_prevmonthhint" : Vals(59) = "Prev Month"
    Keys(60) = "ccs_prevquarterhint" : Vals(60) = "Prev Quarter"
    Keys(61) = "ccs_prevthreemonthshint" : Vals(61) = "Prev Three Months"
    Keys(62) = "ccs_prevyearhint" : Vals(62) = "Prev Year"
    Keys(63) = "ccs_recordformprefix" : Vals(63) = "Add/Edit"
    Keys(64) = "ccs_recordformprefix2" : Vals(64) = "View"
    Keys(65) = "ccs_recordformsuffix" : Vals(65) = ""
    Keys(66) = "ccs_recperpage" : Vals(66) = "Records per page"
    Keys(67) = "ccs_rememberlogin" : Vals(67) = "Remember my Login and Password"
    Keys(68) = "ccs_reportformprefix" : Vals(68) = ""
    Keys(69) = "ccs_reportformsuffix" : Vals(69) = ""
    Keys(70) = "ccs_reportpagenumber1" : Vals(70) = "Page"
    Keys(71) = "ccs_reportpagenumber2" : Vals(71) = "of"
    Keys(72) = "ccs_reportprintlink" : Vals(72) = "Printable version"
    Keys(73) = "ccs_reportsubtotal" : Vals(73) = "Sub Total"
    Keys(74) = "ccs_reporttotal" : Vals(74) = "Grand Total"
    Keys(75) = "ccs_requiredfield" : Vals(75) = "The value in field {0} is required."
    Keys(76) = "ccs_requiredfieldupload" : Vals(76) = "The file attachment in field {0} is required."
    Keys(77) = "ccs_requiredsmtpserver_or_dir" : Vals(77) = "Please specify the SMTP server or Pickup directory for the CDO.Message email component."
    Keys(78) = "ccs_search" : Vals(78) = "Search"
    Keys(79) = "ccs_searchformprefix" : Vals(79) = "Search"
    Keys(80) = "ccs_searchformsuffix" : Vals(80) = ""
    Keys(81) = "ccs_selectfield" : Vals(81) = "Select Field"
    Keys(82) = "ccs_selectorder" : Vals(82) = "Select Order"
    Keys(83) = "ccs_selectvalue" : Vals(83) = "Select Value"
    Keys(84) = "ccs_sortby" : Vals(84) = "Sort by"
    Keys(85) = "ccs_sortdir" : Vals(85) = "Sort direction"
    Keys(86) = "ccs_submitconfirmation" : Vals(86) = "Submit records?"
    Keys(87) = "ccs_tempfoldernotfound" : Vals(87) = "Unable to upload the file specified in {0} - temporary upload folder doesn't exist."
    Keys(88) = "ccs_tempinsufficientpermissions" : Vals(88) = "Insufficient filesystem permissions to upload the file specified in {0} into temporary folder."
    Keys(89) = "ccs_today" : Vals(89) = "Today"
    Keys(90) = "ccs_totalrecords" : Vals(90) = "Total Records:"
    Keys(91) = "ccs_uniquevalue" : Vals(91) = "The value in field {0} is already in database."
    Keys(92) = "ccs_update" : Vals(92) = "Submit"
    Keys(93) = "ccs_uploadcomponenterror" : Vals(93) = "Error occurred while initializing the upload component."
    Keys(94) = "ccs_uploadcomponentnotfound" : Vals(94) = "{0} uploading component {1} is not found. Please install the component or select another one."
    Keys(95) = "ccs_uploadingerror" : Vals(95) = "An error occured when uploading file specified in {0}. Error description: {1}."
    Keys(96) = "ccs_uploadingtempfoldererror" : Vals(96) = "An error occured when uploading file specified in {0} into temporary folder. Error description: {1}."
    Keys(97) = "ccs_wrongtype" : Vals(97) = "The file type specified in field {0} is not allowed."
    CCSLocales.SetKeyVals Keys, Vals  
  End Function
'End CCLoadStaticTranslation

'CCSelectStyle @0-6B1135BA
  Const CCS_SS_RequestParameterName = 0
  Const CCS_SS_CookieName = 1
  Const CCS_SS_SessionName = 2

  Public Sub CCSelectStyle(Path, Default, Names, CookieExpired)
    Dim strStyle : strStyle = Empty
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If Not IsEmpty(Names(CCS_SS_RequestParameterName)) Then
      strStyle = TestStyle(FSO, Path, Request.QueryString(Names(CCS_SS_RequestParameterName)), Default)
      If IsEmpty(strStyle) Then _
        strStyle = TestStyle(FSO, Path, CCGetFromPost(Names(CCS_SS_RequestParameterName),""), Default)
    End If
    If IsEmpty(strStyle) And Not IsEmpty(Names(CCS_SS_CookieName)) Then _
      strStyle = TestStyle(FSO, Path, Request.Cookies(Names(CCS_SS_CookieName)), Default)
    If IsEmpty(strStyle) And Not IsEmpty(Names(CCS_SS_SessionName)) Then _
      strStyle = TestStyle(FSO, Path, Session(Names(CCS_SS_SessionName)), Default)
    If IsEmpty(strStyle) Then _
      strStyle = Default

    If Not IsEmpty(Names(CCS_SS_CookieName)) Then 
       If Request.Cookies(Names(CCS_SS_CookieName)) <> strStyle Then 
         Response.Cookies(Names(CCS_SS_CookieName)) = strStyle
         If Not IsEmpty(CookieExpired) Then _
           Response.Cookies(Names(CCS_SS_CookieName)).Expires = DateAdd("d", CookieExpired, Now())
       End If
    End If
    If Not IsEmpty(Names(CCS_SS_SessionName)) Then _
       Session(Names(CCS_SS_SessionName)) = strStyle
    CCSStyle  = Replace(strStyle," ", "%20")
    Set FSO = Nothing
  End Sub

  Function TestStyle(FSO, Path, Name, Default)
     Dim Res : Res = Empty
     If Len(Name) > 0 Then  
       Name = Trim(Name)
       If CCRegExpTest(Name, "[A-z0-9 ]{1,255}.", True,True) Then
         If FSO.FileExists(Path & Name & "/Style.css") Then 
           Res = Name
         Else 
           Res = Default
         End If
       End If
    End If
    TestStyle = Res
  End Function  

'End CCSelectStyle


%>
