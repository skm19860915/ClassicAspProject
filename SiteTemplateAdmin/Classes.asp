<%
        
'File Description @0-A6CEE862
'======================================================
'
'  This file contains the following classes:
'      Class clsButton
'      Class clsCommand
'      Class clsControl
'      Class clsControls
'      Class clsConverter
'      Class clsDataSource
'      Class clsDatePicker
'      Class clsEmptyDataSource
'      Class clsErrors
'      Class clsEventCaller
'      Class clsField
'      Class clsFields
'      Class clsFileElement
'      Class clsFileUpload
'      Class clsFormElement
'      Class clsListControl
'      Class clsReportTotalRef
'      Class clsLocales
'      Class clsSQLParameter
'      Class clsSQLParameters
'      Class clsSection
'      Class clsStringBuffer
'      Class clsUploadControl
'
'======================================================
'End File Description

'Constant List @0-B818DB51

' ------- Controls ---------------
Const ccsLabel             = 00001
Const ccsLink              = 00002
Const ccsTextBox           = 00003
Const ccsTextArea          = 00004
Const ccsListBox           = 00005
Const ccsRadioButton       = 00006
Const ccsButton            = 00007
Const ccsCheckBox          = 00008
Const ccsImage             = 00009
Const ccsImageLink         = 00010
Const ccsHidden            = 00011
Const ccsCheckBoxList      = 00012
Const ccsDatePicker        = 00013
Const ccsReportLabel       = 00014
Const ccsPageBreak         = 00015

Dim ccsControlTypes(15)
ccsControlTypes(ccsLabel)        = "Label"
ccsControlTypes(ccsReportLabel)  = "ReportLabel"
ccsControlTypes(ccsPageBreak)    = "PageBreak"
ccsControlTypes(ccsLink)         = "Link"
ccsControlTypes(ccsTextBox)      = "TextBox"
ccsControlTypes(ccsTextArea)     = "TextArea"
ccsControlTypes(ccsListBox)      = "ListBox"
ccsControlTypes(ccsRadioButton)  = "RadioButton"
ccsControlTypes(ccsButton)       = "Button"
ccsControlTypes(ccsCheckBox)     = "CheckBox"
ccsControlTypes(ccsImage)        = "Image"
ccsControlTypes(ccsImageLink)    = "ImageLink"
ccsControlTypes(ccsHidden)       = "Hidden"
ccsControlTypes(ccsCheckBoxList) = "CheckBoxList"
ccsControlTypes(ccsDatePicker)   = "DatePicker"

' ------- Operators --------------
Const opEqual              = 00001
Const opNotEqual           = 00002
Const opLessThan           = 00003
Const opLessThanOrEqual    = 00004
Const opGreaterThan        = 00005
Const opGreaterThanOrEqual = 00006
Const opBeginsWith         = 00007
Const opNotBeginsWith      = 00008
Const opEndsWith           = 00009
Const opNotEndsWith        = 00010
Const opContains           = 00011
Const opNotContains        = 00012
Const opIsNull             = 00013
Const opNotNull            = 00014

' ------- Datasource types -------
Const dsTable              = 00001
Const dsSQL                = 00002
Const dsProcedure          = 00003
Const dsListOfValues       = 00004
Const dsEmpty              = 00005

' ------- Command types -------
Const cmdOpen              = 00001
Const cmdExec              = 00002

' ------- Parse types ------------
Const ccsParseAccumulate   = True
Const ccsParseOverwrite    = False

' ------- Listbox populating types ------------
Const ccsJoins             = 0
Const ccsStringConcats     = 1

' ------- CheckBox states --------
Const ccsChecked           = True
Const ccsUnchecked         = False
'End Constant List

'clsSQLParameters Class @0-7B9EE410

Class clsSQLParameters
  
  Public Connection
  Public Criterion()
  Public AssembledWhere
  Public ParameterSources
  Public Errors
  Public DataSource

  Public ParametersList

  Private Sub Class_Initialize()
    ReDim Criterion(100)
    Set ParametersList = Server.CreateObject("Scripting.Dictionary")
    Set DataSource = Nothing
  End Sub

  Private Sub Class_Terminate()
    Set ParametersList = Nothing
  End Sub

  Public Default Property Get Parameters(Name)
    Set Parameters = ParametersList(Name)
  End Property

  Public Property Set Parameters(Name, NewParameter)
    Set ParametersList(Name) = NewParameter
  End Property

  Property Get Count()
    Count = ParametersList.Count
  End Property

  Function AddParameter(ID, ParameterSource, DataType, Format, DBFormat, DefaultValue, UseIsNull)
    Dim SQLParameter
    Set SQLParameter = New clsSQLParameter
    With SQLParameter
      Set .Connection = Connection
      .DataType = DataType
      .Format = Format
      .DBFormat = DBFormat
      .Caption = ParameterSource
      .DefaultValue = DefaultValue
      .UseIsNull = UseIsNull
      Set .DataSource = DataSource
      If IsObject(ParameterSources) Then
        If IsEmpty(DefaultValue) Or ((DataType <> ccsText And DataType <> ccsMemo) And CStr(DefaultValue) = "") Then
          .Text = ParameterSources(ParameterSource)
        Else
          If IsEmpty(ParameterSources(ParameterSource)) Or ((DataType <> ccsText And DataType <> ccsMemo) And ParameterSources(ParameterSource) = "") Then
            .Text = DefaultValue
          Else
            .Text = ParameterSources(ParameterSource)
          End If
        End If
      End If
    End With
    Set ParametersList(ID) = SQLParameter
    Set SQLParameter = Nothing
  End Function

  Function getParamByID(ID)
    Set getParamByID = ParametersList(ID)
  End Function

  Property Get AllParamsSet()
    Dim ParametersItems, I, Result

    Result = True
    I = 0
    ParametersItems = ParametersList.Items
    While Result AND (I <= UBound(ParametersItems))
      Result = (NOT IsEmpty(ParametersItems(I).Value)) OR (IsEmpty(ParametersItems(I).Value) AND ParametersItems(I).UseIsNull)
      I = I + 1
    Wend
    AllParamsSet = Result
  End Property

  Function GetError()
    Dim ParametersItems, I, Result

    ParametersItems = ParametersList.Items
    For I = 0 To UBound(ParametersItems)
      Result = Result & ParametersItems(I).Errors.ToString
    Next
    GetError = Result
  End Function

  Function opAND(Brackets, LeftPart, RightPart)
    Dim Result
    If NOT IsEmpty(LeftPart) Then
      If NOT IsEmpty(RightPart) Then
        Result = LeftPart & " and " & RightPart
      Else
        Result = LeftPart
      End If
    Else
      If NOT IsEmpty(RightPart) Then
        Result = RightPart
      End If
    End If
    If Brackets And NOT IsEmpty(Result) Then _
      Result = " (" & Result & ") "
    opAND = Result
  End Function

  Function opOR(Brackets, LeftPart, RightPart)
    Dim Result
    If NOT IsEmpty(LeftPart) Then
      If NOT IsEmpty(RightPart) Then
        Result = LeftPart & " or " & RightPart
      Else
        Result = LeftPart
      End If
    Else
      If NOT IsEmpty(RightPart) Then
        Result = RightPart
      End If
    End If
    If Brackets And NOT IsEmpty(Result) Then _
        If Brackets Then Result = " (" & Result & ") "
    opOR = Result
  End Function

  Function Operation(Operator, Brackets, FieldName, Parameter)
    If CStr(Parameter.Text) <> "" Then
      Dim Result
      Dim Value, SQLValue
      Value = IIF( Parameter.DataType = ccsText OR Parameter.DataType = ccsMemo, Parameter.Value, Parameter.SQLText)
      SQLValue = Connection.ToSQL(Value, Parameter.DataType)

      Select Case Operator
        Case opEqual
          Result = FieldName & " = " & SQLValue
        Case opNotEqual
          Result = FieldName & " <> " & SQLValue
        Case opLessThan
          Result = FieldName & " < " & SQLValue
        Case opLessThanOrEqual
          Result = FieldName & " <= " & SQLValue
        Case opGreaterThan
          Result = FieldName & " > " & SQLValue
        Case opGreaterThanOrEqual
          Result = FieldName & " >= " & SQLValue
        Case opBeginsWith, opNotBeginsWith, opEndsWith, opNotEndsWith, opContains, opNotContains
          Result = FieldName & Connection.ToLikeCriteria(Connection.EscapeChars(Value), Operator)
        Case opIsNull
          Result = FieldName & " is null"
        Case opNotNull
          Result = FieldName & " is not null"
      End Select
      Operation = Result
    Else
     If Parameter.UseIsNull Then
       Select Case Operator
         Case opNotEqual, opNotBeginsWith, opNotEndsWith, opNotContains, opNotNull
           Result = FieldName & " is not null"
         Case Else
           Result = FieldName & " is null"
       End Select
       Operation = Result
     Else
       Operation = Empty
     End If
    End If
  End Function
End Class
'End clsSQLParameters Class

'clsSQLParameter Class @0-29CF4C98
Class clsSQLParameter
  Public Errors
  Public DataType
  Public Format
  Public DBFormat
  Public Caption
  Public Connection
  Public DataSource
  Public DefaultValue
  Public UseIsNull

  Private VarValue
  Private SQLTextValue
  Private TextValue

  Private Sub Class_Initialize()
    VarValue = Empty
    SQLTextValue = Empty
    TextValue = Empty
    UseIsNull = False

    DataType = ccsText
    Set Errors = New clsErrors
  End Sub
  
  Private Sub Class_Terminate()
    Set Errors = Nothing
  End Sub

  Function GetParsedValue(ParsingValue, MaskFormat)
    Dim Result

    If Not IsEmpty(ParsingValue) Then
      Select Case DataType
        Case ccsDate
          If VarType(ParsingValue) = vbDate Then
            Result = ParsingValue
          ElseIf CCValidateDate(ParsingValue, MaskFormat) Then
            Result = CCParseDate(ParsingValue, MaskFormat)
          Else
            If IsArray(Format) Then 
              PrintDBError "", "", CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(Format)))
            Else 
              PrintDBError "", "", CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(Format)))
            End If
          End If
        Case ccsBoolean
          Result = CCParseBoolean(ParsingValue, MaskFormat)
        Case ccsInteger
          If CCValidateNumber(ParsingValue, MaskFormat) Then 
            Result = CCParseInteger(ParsingValue, MaskFormat) 
          Else 
            PrintDBError "", "", CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(Format)))
          End If
        Case ccsFloat
          If CCValidateNumber(ParsingValue, MaskFormat) Then 
            Result = CCParseFloat(ParsingValue, MaskFormat)
          Else
            PrintDBError "", "", CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(Format)))
          End If
        Case ccsSingle
          If CCValidateNumber(ParsingValue, MaskFormat) Then 
            Result = CCParseSingle(ParsingValue, MaskFormat)
          Else
            PrintDBError "", "", CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(Format)))
          End If
        Case ccsText, ccsMemo
          Result = CStr(ParsingValue)
      End Select
    End If

    GetParsedValue = Result
  End Function

  Function GetFormattedValue(MaskFormat)
    Dim Result, Value

    If IsEmpty(VarValue) Then 
      Value = DefaultValue 
    Else 
      Value = VarValue
    End If

    Select Case DataType
      Case ccsDate
        Result = CCFormatDate(Value, MaskFormat)
      Case ccsBoolean
        Result = CCFormatBoolean(Value, MaskFormat)
      Case ccsInteger, ccsFloat, ccsSingle
        Result = CCFormatNumber(Value, MaskFormat)
      Case ccsText, ccsMemo
        Result = CStr(Value)
        If CStr(Result) <> "" Then Result = Connection.EscapeChars(Result)
    End Select

    GetFormattedValue = Result
  End Function

  Property Let Value(NewValue)
    VarValue = Empty
    SQLTextValue = Empty
    TextValue = Empty
    If NOT IsEmpty(NewValue) And Not (NewValue="") Then
      Select Case DataType
        Case ccsDate
          VarValue = CDate(NewValue)
        Case ccsBoolean
          VarValue = CBool(NewValue)
        Case ccsInteger
          VarValue = CLng(NewValue)
        Case ccsFloat
          VarValue = CDbl(NewValue)
        Case ccsSingle
          VarValue = CSng(NewValue)
        Case ccsText, ccsMemo
          VarValue = CStr(NewValue)
      End Select
    End If
  End Property

  Property Get Value()
    If IsEmpty(VarValue) Then 
      Value = DefaultValue 
    Else
      Value = VarValue
    End If
  End Property

  Property Let Text(NewText)
    If Not IsEmpty(NewText) Then
      SQLTextValue = Empty
      TextValue = NewText
      VarValue = GetParsedValue(TextValue, Format)
    End If
  End Property

  Property Get Text()
    If IsEmpty(TextValue) Then TextValue = GetFormattedValue(Format)
    Text = TextValue
  End Property

  Property Let SQLText(varNewSQLText)
    SQLTextValue = varNewSQLText
  End Property

  Property Get SQLText()
    If IsEmpty(SQLTextValue) Then 
      SQLTextValue = GetFormattedValue(DBFormat)
    End If
    SQLText = SQLTextValue
  End Property

End Class

'End clsSQLParameter Class

'clsFields Class @0-791D3D1C
Class clsFields
  Private objFields
  Private Items
  Private Counter

  Private Sub Class_Initialize()
    Set objFields = CreateObject("Scripting.Dictionary")
  End Sub

  Sub AddFields(Fields) ' Add new objects to Object array
    Dim I
    If IsArray(Fields) Then
      For I = LBound(Fields) To UBound(Fields)
        Set objFields(Fields(I).Name) = Fields(I)
      Next
    End If
  End Sub

  Public Default Property Get Item(Name)
    Set Item = objFields(Name)
  End Property

  Sub InitEnum()
    Items = objFields.Items
    Counter = 0
  End Sub

  Function NextItem()
    Set NextItem = Items(Counter)
    Counter = Counter + 1
  End Function

  Function EndOfEnum()
    EndOfEnum = (Counter > UBound(Items))
  End Function

  Function Exists(Name)
    Exists = objFields.Exists(Name)
  End Function

End Class
'End clsFields Class

'CCCreateCollection Function @0-61899D71
Function CCCreateCollection(Block, TargetBlock, Accumulate, Controls)
  Dim Collection
  Set Collection = New clsControls
  With Collection
    Set .Block = Block
    If NOT IsNull(TargetBlock) Then
      Set .TargetBlock = TargetBlock
    End If
    .Accumulate = Accumulate
    .AddControls Controls
  End With
  Set CCCreateCollection = Collection
End Function

'End CCCreateCollection Function

'clsControls Class @0-64CB7C37
Class clsControls
  Private Objects ' Dictionary object
  Private CCSEventResult
  Private EnumData
  Private Counter
  
  Public Block
  Public Accumulate

  Private objTargetBlock
  Private isSetTargetBlock
  Private mVisible

  Private Sub Class_Initialize()
    Set Objects = Server.CreateObject("Scripting.Dictionary")
    mVisible = True
  End Sub

  Private Sub Class_Terminate()
    Set Objects = Nothing
  End Sub

  Sub AddControls(Controls) ' Add new objects to Object array  
    Dim ArraySize, NumberControls, I

    If IsArray(Controls) Then
      NumberControls = UBound(Controls)
      ArraySize = Objects.Count

      For i = ArraySize To ArraySize + NumberControls
        Objects.Add i,Controls(I)
      Next
    End If
  End Sub

  Sub AddControl(Control) ' Add a new object to Object array 
    If TypeName(Control) = "clsControls" Then 
      Objects.Add Objects.Count, Control
    Else
      Objects.Add Control.Name, Control
    End If
  End Sub

  Property Get Items(ItemName)
    If Objects.Exists(ItemName) Then
      Set Items = Objects(ItemName)
    Else
      Set Items = Nothing
    End If
  End Property
  
  Property Let Items(ItemName, NewItem)
    If Objects.Exists(ItemName) Then
      Objects(ItemName) = NewItem
    Else
      Objects.Add ItemName, NewItem
    End If
  End Property

  Function GetItemByName(ItemName)
    Dim Element
    For Each Element In Objects
      If Objects.Item(Element).Name = ItemName Then
        Set GetItemByName = Objects.Item(Element)
        Exit Function
      End If
    Next
    Set GetItemByName = Nothing
  End Function

  Sub Show()
    Dim Element, Obj

    If NOT mVisible Then Exit Sub

    For Each Element In Objects
      Set Obj = Objects.Item(element)
      If TypeName(Obj) = "clsControls" Then
        Obj.Show
      Else
        Obj.Show Block 
      End If
    Next

    If Not IsEmpty(Accumulate) Then
      If isSetTargetBlock Then
        Block.ParseTo Accumulate, objTargetBlock
      Else
        Block.Parse Accumulate
      End If
    End If

  End Sub


  Sub Validate()
    Dim Element
    For Each Element In Objects
      Objects.Item(Element).Validate
    Next
  End Sub

  Function isValid()
    Dim Element
    For Each Element In Objects
      If Objects.Item(Element).Errors.Count > 0 Then
        isValid = False
        Exit Function
      End If
    Next
    isValid = True
  End Function

  Function GetErrors()
    Dim Errors, Element
    For Each Element In Objects
      Errors = Errors & Objects.Item(Element).Errors.ToString
    Next
    GetErrors = Errors
  End Function

  Function GetErrorsArray()
    Dim Errors,Element
    Set Errors = New clsErrors
    If Objects.Count>0 Then 
     For Each Element In Objects
      Errors.AddErrors Objects.Item(Element).Errors
     Next
    End If
    Set GetErrorsArray=Errors
  End Function


  Property Set TargetBlock(NewBlock)
    isSetTargetBlock = True
    Set objTargetBlock = NewBlock
  End Property

  Sub InitEnum()
    EnumData = Objects.Items
    Counter = 0
  End Sub
  
  Function NextItem()
    Set NextItem = EnumData(Counter)
    Counter = Counter + 1
  End Function
  
  Function EndOfEnum()
    EndOfEnum = (Counter > UBound(EnumData))
  End Function 

  Property Let Visible(newValue)
    mVisible = CBool(newValue)
  End Property

  Property Get Visible()
    Visible = mVisible
  End Property

  Sub PreserveControlsVisible()
    Dim  Element, Obj 
    For Each Element In Objects
      Set Obj = Objects.Item(Element)
      If TypeName(Obj) = "clsPanel" Then
        Obj.PreserveControlsVisible 
      ElseIf TypeName(Obj) = "clsControl" Or TypeName(Obj) = "clsListControl" Then
        Obj.PreserveVisible = Obj.Visible
      End If
    Next
  End Sub

  Sub RestoreControlsVisible()
    Dim  Element, Obj 
    For Each Element In Objects
      Set Obj = Objects.Item(Element)
      If TypeName(Obj) = "clsPanel" Then
        Obj.RestoreControlsVisible 
      ElseIf TypeName(Obj) = "clsControl" Or TypeName(Obj) = "clsListControl" Then
        Obj.Visible = Obj.PreserveVisible
      End If
    Next
  End Sub

End Class

'End clsControls Class

'CCCreateControl Function @0-706BB4D6
Function CCCreateControl(ControlType,  Name, Caption, DataType, Format, InitValue)
  Dim Control
  Set Control = New clsControl
  With Control    
    .ControlType = ControlType
    .Name = Name
    .BlockName = ccsControlTypes(ControlType) & " " & Name
    .ControlTypeName = ccsControlTypes(ControlType)
    .Caption = Caption
    .DataType = DataType
    .Format = Format
    If NOT IsEmpty(InitValue) Then
      If ControlType = ccsCheckBox Then
          .State = True
      Else
        .Text = InitValue
      End If
    End If
  End With
  Set CCCreateControl = Control
End Function
'End CCCreateControl Function

'CCCreateReportLabel Function @0-9BB6254F
Function CCCreateReportLabel(Name, Caption, DataType, Format, InitValue, TotalFunction, IsPercent, IsEmptySource,EmptyText)
  Dim Control
  Set Control = CCCreateControl(ccsReportLabel,  Name, Caption, DataType, Format, InitValue)
  With Control    
    .ControlType = ccsReportLabel
    .ControlTypeName=ccsControlTypes(ccsReportLabel)
    .TotalFunction = TotalFunction
    .IsPercent = IsPercent
    .IsEmptySource = IsEmptySource
  .EmptyText = EmptyText
  End With
  Set CCCreateReportLabel = Control
End Function
'End CCCreateReportLabel Function

'clsControl Class @0-45AADF54
Class clsControl
  Public Errors
  Public DataType
  Public Format
  Public DBFormat
  Public ControlType
  Public ControlTypeName
  Public Name
  Public BlockName
  Public ExternalName
  Public HTML
  Public Required
  Public CheckedValue
  Public UncheckedValue
  Public State
  Public Visible
  Public TemplateBlock
  Public InputMask
  Public CountValue
  Public SumValue
  Public ValueRelative
  Public CountValueRelative
  Public SumValueRelative
  Public TotalFunction
  Public IsPercent
  Public IsEmptySource
  Public PreserveVisible

  Public Parameters

  Private isInternal
  Public  initialValue
  Private prevItem
  Private prevValue
  Private prevCountValue
  Private prevSumValue
  Private prevValueRelative
  Private prevCountValueRelative
  Private prevSumValueRelative

  Private mPage
  Private VarValue
  Private TextValue
  Private m_EmptyText
  Private mCaption

  Public CCSEvents
  Private CCSEventResult

  Private Sub Class_Initialize()
    VarValue = Empty
    TextValue = Empty
    m_EmptyText = Empty
    Visible = True
    ExternalName = Empty
    CheckedValue = True
    UncheckedValue = False
  IsEmptySource = False
  initialValue = Empty
  prevItem = False
  isInternal = False
    DataType = ccsText
    HTML = False
    Required = False
    Set Errors = New clsErrors
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Parameters = ""
  End Sub

  Private Sub Class_Terminate()
    Set Errors = Nothing
    Set CCSEvents = Nothing
  End Sub

  Function Validate()
    If Required And CStr(VarValue) = "" And Errors.Count = 0 Then
      Errors.addError(CCSLocales.GetText("CCS_RequiredField", Me.Caption))
    End If
    If Errors.Count = 0 AND Not IsEmpty(InputMask) Then _
      If NOT CCRegExpTest(Me.Text, InputMask, False, False) Then _
        Errors.addError(CCSLocales.GetText("CCS_MaskValidation", Me.Caption))
    Validate = CCRaiseEvent(CCSEvents, "OnValidate", Me)
  End Function

  Function GetParsedValue(ParsingValue, MaskFormat)
    Dim Result

    Result = CCSConverter.StringToType(DataType, ParsingValue, MaskFormat)
    If CCSConverter.ParseError Then
      If DataType = ccsDate AND (IsArray(Format) OR IsArray(CCSConverter.DateFormat)) Then
        If IsArray(Format) Then _
          Errors.addError(CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(Format)))) _
        Else _
          Errors.addError(CCSLocales.GetText("CCS_IncorrectFormat", Array(Caption, CCGetFormatStr(CCSConverter.DateFormat))))
      Else
        Errors.addError(CCSLocales.GetText("CCS_IncorrectValue", Array(Caption)))
      End If
    End If

    GetParsedValue = Result
  End Function

  Public Sub SetPercentValue(val, relativeVal)
    If IsEmpty(relativeVal) Or IsEmpty(val) Then Value = Empty:Exit Sub
    If relativeVal = 0 Then Value = Empty:Exit Sub
    Value = val / relativeVal
  End Sub

  Public Function GetTotalValue(isCalculate)
    If Not isCalculate Then
      isInternal = True
      If TotalFunction = "Count" And  IsEmpty(prevValue) Then prevValue = 0
      Value = prevValue
      isInternal = False
      GetTotalValue = prevValue:Exit Function
    End If
    isInternal = True
    Value = initialValue
    isInternal = False
    Dim newVal
  Select Case TotalFunction
      Case "Sum"
        If IsEmpty(Value) And IsEmpty(prevValue) Then
          newVal = Empty
        Else
          If DataType = ccsSingle Then
            newVal = CSng(Value) + CSng(prevValue)
          Else
            newVal = CDbl(Value) + CDbl(prevValue)
          End If
        End If
        If IsPercent And IsEmpty(Value) And IsEmpty(prevValueRelative) Then
          ValueRelative = Empty
        ElseIf IsPercent Then
          If DataType = ccsSingle Then
            ValueRelative = CSng(Value) + CSng(prevValueRelative)
          Else
            ValueRelative = CDbl(Value) + CDbl(prevValueRelative)
          End If
        End If
      Case ""
    newVal = Value
  If IsPercent And IsEmpty(Value) And IsEmpty(prevValueRelative) Then
          ValueRelative = Empty
        ElseIf IsPercent Then
          If DataType = ccsSingle Then
            ValueRelative = CSng(Value) + CSng(prevValueRelative)
          Else
            ValueRelative = CDbl(Value) + CDbl(prevValueRelative)
          End If
        End If
      Case "Count"
        If DataType = ccsSingle Then
          newVal = CSng(prevValue) + IIf(IsEmptySource,1,Abs(CSng(Not IsEmpty(Value))))
          If IsPercent Then ValueRelative = CSng(prevValueRelative) + IIf(IsEmptySource,1,Abs(CSng(Not IsEmpty(Value))))
        Else
          newVal = CDbl(prevValue) + IIf(IsEmptySource,1,Abs(CDbl(Not IsEmpty(Value))))
          If IsPercent Then ValueRelative = CDbl(prevValueRelative) + IIf(IsEmptySource,1,Abs(CDbl(Not IsEmpty(Value))))
        End If

      Case "Min"
        newVal = Min(Value,prevValue)
        If IsPercent Then ValueRelative = Min(Value, prevValueRelative)
      Case "Max"
        newVal = Max(Value,prevValue)
        If IsPercent Then ValueRelative = Max(Value, prevValueRelative)
      Case "Avg"
        If Not IsEmpty(Value) Then 
          CountValue = prevCountValue + 1
          If DataType = ccsSingle Then
            SumValue = CSng(prevSumValue) + CSng(Value)
          Else
            SumValue = CDbl(prevSumValue) + CDbl(Value)
          End If
        End If
        If CountValue=0 Then 
          newVal = prevValue
        Else
          newVal = SumValue / CountValue
        End if
        If IsPercent Then 
          If Not IsEmpty(Value) Then 
            CountValueRelative = prevCountValueRelative + 1
            If DataType = ccsSingle Then
              SumValueRelative = CSng(prevSumValueRelative) + CSng(Value)
            Else
              SumValueRelative = CDbl(prevSumValueRelative) + CDbl(Value)
            End If
          End If
          If CountValueRelative=0 Then 
            ValueRelative = prevValueRelative
          Else
            ValueRelative = SumValueRelative / CountValueRelative
          End if
        End If
    End Select
    prevValueRelative = ValueRelative
    prevValue = newVal
    prevCountValue = CountValue
    prevSumValue = SumValue
    prevCountValueRelative = CountValueRelative
    prevSumValueRelative = SumValueRelative
    isInternal = True
    Value = newVal
    isInternal = False
    GetTotalValue = Value
  End Function

  Public Sub Reset()
    prevValue = Empty
    prevCountValue = Empty
    prevSumValue = Empty
    sumValue=Empty
    CountValue = Empty
  End Sub

  Public Sub ResetRelativeValues()
    ValueRelative = initialValue
    prevValueRelative = Empty
    prevCountValueRelative = Empty
    prevSumValueRelative = Empty
  End Sub

  Property Get Link()
    Dim Result

    If Parameters = "" Then
      Result = mPage
    Else
      Result = mPage & "?" & Parameters
    End If

    Link = Result
  End Property

  Property Let Link(newLink)
    Dim parsedLink
    If CStr(newLink) = "" Then
      mPage = ""
      Parameters = ""
    Else
      parsedLink = Split(newLink, "?")
      mPage = parsedLink(0)
      If UBound(parsedLink) = 1 Then 
        Parameters = parsedLink(1) 
      Else
        Parameters = ""
      End If
    End If
  End Property

  Property Get Page()
    Page = mPage
  End Property

  Property Let Page(newPage)
    mPage = newPage
  End Property

  Property Get Caption()
    If Len(mCaption) > 0 Then
      Caption = mCaption
    Else
      Caption = Name
    End If
  End Property

  Property Let Caption(newValue)
    mCaption = newValue
  End Property

  Function GetFormattedValue(MaskFormat)
    If ControlType = ccsReportLabel And Not IsEmpty(VarValue) And (DataType=ccsText Or DataType=ccsFloat Or DataType=ccsSingle) And IsArray(MaskFormat) Then 
      Select Case DataType
        Case ccsFloat
           GetFormattedValue = CCSConverter.TypeToString(ccsFloat,CDbl(VarValue), MaskFormat)
        Case ccsSingle
           GetFormattedValue = CCSConverter.TypeToString(ccsSingle,CSng(VarValue), MaskFormat)
        Case Else
           GetFormattedValue = CCSConverter.TypeToString(ccsFloat,CDbl(VarValue), MaskFormat)
      End Select 
    Else     
      GetFormattedValue = CCSConverter.TypeToString(DataType, VarValue, MaskFormat)
    End If
  End Function

  Sub Show(Template)
    Dim NeedShow, sTmpValue

    Set TemplateBlock = Template.Block(ControlTypeName & " " & Name)

    If TemplateBlock Is Nothing Then
      Set TemplateBlock = Template
      NeedShow = False
    Else
      NeedShow = True
      TemplateBlock.HTML = ""
    End If
    
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
    If NOT Visible Then
      TemplateBlock.Variable(Name)=""
      Exit Sub
    End If

    If IsEmpty(ExternalName) Then
      TemplateBlock.Variable(Name & "_Name") = Name
    Else
      TemplateBlock.Variable(Name & "_Name") = ExternalName
    End If

    If IsEmpty(TextValue) Then 
      TextValue = GetFormattedValue(Format)
    End If

    Select Case ControlType
      Case ccsPageBreak
          TemplateBlock.Variable(Name) = TextValue
      Case ccsLabel, ccsReportLabel, ccsTextBox, ccsTextArea, ccsHidden
        If TextValue = "" Or IsEmpty(TextValue) Then TextValue = EmptyText
        If HTML Then
          TemplateBlock.Variable(Name) = TextValue
        Else
          sTmpValue = Server.HTMLEncode(TextValue)
          If ControlType = ccsLabel or ControlType = ccsReportLabel Then 
            sTmpValue = Replace(sTmpValue, vbCrLf, "<BR>")
          End If
          TemplateBlock.Variable(Name) = sTmpValue
        End If
      Case ccsImage
          sTmpValue = Server.HTMLEncode(TextValue)
          If ControlType = ccsLabel or ControlType = ccsReportLabel Then 
            sTmpValue = Replace(sTmpValue, vbCrLf, "<BR>")
          End If
          TemplateBlock.Variable(Name) = sTmpValue
      Case ccsLink
        If HTML Then 
          TemplateBlock.Variable(Name) = TextValue 
        Else 
          TemplateBlock.Variable(Name) = Replace(Server.HTMLEncode(TextValue), vbCrLf, "<BR>")
        End If
        TemplateBlock.Variable(Name & "_Src") = Me.Link
      Case ccsImageLink
        TemplateBlock.Variable(Name & "_Src") = Server.HTMLEncode(TextValue) 
        TemplateBlock.Variable(Name) = Me.Link
      Case ccsCheckBox
        If State Then 
          TemplateBlock.Variable(Name) = "CHECKED" 
        Else
          TemplateBlock.Variable(Name) = ""
        End If
    End Select
    If NeedShow Then TemplateBlock.Show
    Set TemplateBlock = Nothing
  End Sub

  Property Let Value(NewValue)
    VarValue = Empty
    TextValue = Empty
    VarValue = CCSConverter.VBSConvert(DataType, NewValue)

    If ControlType = ccsCheckBox Then
      If DataType = ccsBoolean Then
        If IsEmpty(NewValue) Or (NewValue="") Then
          State = False
        Else
          State = VarValue
        End If
      Else
        if DataType = ccsDate Then
          State = (VarValue = CDate(CheckedValue))
        Else
          State = (VarValue = CheckedValue)
        End if
      End If
    End If
  If Not isInternal Then initialValue = VarValue
  End Property

  Property Get Value()
    If ControlType = ccsCheckBox Then
      If IsEmpty(State) Then 
        Value = UncheckedValue 
      Else 
        Value = IIf(State, CheckedValue, UncheckedValue)
      End If
    Else
        Value = VarValue
    End If
  End Property

  Property Let Text(NewText)
    VarValue = Empty
    TextValue = NewText
    If ControlType = ccsCheckBox Then
      VarValue = IIf(IsEmpty(NewText), UncheckedValue, CheckedValue)
      State = (VarValue = CheckedValue)
    Else 
      VarValue = GetParsedValue(TextValue, Format)
    End If
  End Property

  Property Get Text()
    If IsEmpty(TextValue) Then 
      TextValue = GetFormattedValue(Format)
    End If
    Text = TextValue
  End Property

  Property Let EmptyText(NewText)
    m_EmptyText = NewText
  End Property

  Property Get EmptyText()
    EmptyText = m_EmptyText
  End Property

  Function ChangeValue(NewValue)
   VarValue=NewValue
   TextValue = Empty
  End Function

End Class
'End clsControl Class

'CCCreateField Function @0-A187BD87
Function CCCreateField(Name, DBFieldName, DataType, DBFormat, DataSource)
  Dim Field
  Set Field = New clsField
  With Field
    .Name = Name
    .DBFieldName = DBFieldName
    .DataType = DataType
    .DBFormat = DBFormat
    Set .DataSource = DataSource
  End With
  Set CCCreateField = Field
End Function
'End CCCreateField Function

'clsField Class @0-35CA46D8
Class clsField
  Public DataType
  Public DBFormat
  Public Name
  Public DBFieldName
  Public Errors

  Private mDataSource
  Private mConverter
  Private mValue
  Private mSQLText

  Private Sub Class_Initialize()
    mValue = Empty
    mSQLText = Empty
    DataType = ccsText
    Set Errors = New clsErrors
  End Sub
  
  Private Sub Class_Terminate()
    Set Errors = Nothing
  End Sub

  Public Function GetParsedValue(ParsingValue, MaskFormat)
    Dim Result, ValueType
    Result = Empty

    If NOT IsEmpty(ParsingValue) Then
      ValueType = VarType(ParsingValue)
      If ValueType = vbString Or (DataType=ccsBoolean And (ValueType>=2 And ValueType<=5)) Then
        Result = mConverter.StringToType(DataType, ParsingValue, MaskFormat)
      Else
        Result = mConverter.VBSConvert(DataType, ParsingValue)
      End If
      If mConverter.ParseError Then
        Errors.addError(CCSLocales.GetText("CCS_IncorrectValue", Array(Name)))
      End If
    End If
    GetParsedValue = Result
  End Function

  Public Function GetFormattedValue(MaskFormat)
    Dim Result
    Result = CCSConverter.TypeToString(DataType, mValue, MaskFormat)
    Select Case DataType
      Case ccsText, ccsMemo
        If CStr(Result) <> "" Then Result = mDataSource.DataSource.Connection.EscapeChars(Result)
    End Select
    GetFormattedValue = Result
  End Function

  Public Property Let Value(vData)
    mSQLText = Empty
    mValue = mConverter.VBSConvert(DataType, vData)
  End Property

  Public Default Property Get Value()
    Dim RS, Result, ResultType

    Result = mValue
    If IsObject(mDataSource.Recordset) Then
      Set RS = mDataSource.Recordset
      If RS.State = adStateOpen Then
        If NOT RS.EOF Then _
          Result = CCGetValue(RS, DBFieldName) 
      End If
      Result = GetParsedValue(Result, DBFormat)
    End If
    Value = Result
  End Property

  Public Property Let SQLText(vData)
    mSQLText = vData
  End Property

  Public Property Get SQLText()
    If IsEmpty(mSQLText) Then 
      mSQLText = GetFormattedValue(DBFormat)
    End If
    SQLText = mSQLText
  End Property
  
  Public Property Set DataSource(oRef)
    Set mDataSource = oRef
    If IsObject(mDataSource.Connection) Then
      Set mConverter = mDataSource.Connection.Converter
    Else
      Set mConverter = CCSConverter
    End If
  End Property

  Public Property Get DataSource()
    Set DataSource = mDataSource
  End Property

End Class
'End clsField Class

'CCCreatePanel Function @0-BA92F9B4
Function CCCreatePanel(Name)
    Dim Panel
    Set Panel = New clsPanel
    Panel.Name = Name
    Set CCCreatePanel = Panel
End Function 

'End CCCreatePanel Function

'clsPanel Class @0-0AFD74C6
class clsPanel
  Public Name
  Public CCSEvents
  Public Visible
  Public ExternalName
  Public Components
  Public PreserveVisible

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set Components = CreateObject("Scripting.Dictionary")
    ExternalName = Empty
    Visible = True
  End Sub

  Private Sub Class_Terminate()
    Set CCSEvents = Nothing
    Set Components = Nothing
  End Sub

  
  Public Function AddComponent(Component)
    Components.Add GetComponentName(Component), Component
  End Function

  Public Function AddComponents(Components)
    Dim I
    If IsArray(Components) Then 
      For I = 0 To UBound(Components)
      Me.Components.Add GetComponentName(Components(I)), Components(I)
      Next
    Else 
      Me.Components.Add GetComponentName(Components(I)), Components
    End If
  End Function

  Private Function GetComponentName(Component)
    On Error Resume Next
    GetComponentName = Component.Name
    If Err.Number <> 0 Then
      GetComponentName = Component.ComponentName
    End If
    On Error Goto 0
  End Function

  Sub Show(Template)
    Dim Key, Obj, PanelBlock, BlockName
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
    BlockName = "Panel " & Name
    If Visible Then
      If Template.BlockExists(BlockName, "block") Then 
        Set PanelBlock = Template.Block(BlockName)
        For Each Key in Components
          Set Obj = Components(key)
          Obj.Show PanelBlock
        Next
        PanelBlock.Parse ccsParseOverwrite
      End If
    Else
      Template.Template.HideBlock(BlockName)
    End If
  End Sub

  Sub PreserveControlsVisible()
    Dim  Element, Obj 
    PreserveVisible = Visible
    For Each Element In Components
      Set Obj = Components.Item(Element)
      If TypeName(Obj) = "clsPanel" Then
        Obj.PreserveControlsVisible 
      ElseIf TypeName(Obj) = "clsControl" Or TypeName(Obj) = "clsListControl" Then
        Obj.PreserveVisible = Obj.Visible
      End If
    Next
  End Sub

  Sub RestoreControlsVisible()
    Dim  Element, Obj 
    Visible = PreserveVisible
    For Each Element In Components
      Set Obj = Components.Item(Element)
      If TypeName(Obj) = "clsPanel" Then
        Obj.RestoreControlsVisible 
      ElseIf TypeName(Obj) = "clsControl" Or TypeName(Obj) = "clsListControl" Then
        Obj.Visible = Obj.PreserveVisible
      End If
    Next
  End Sub

End Class

'End clsPanel Class

'CCCreateCalendarNavigator Function @0-289CB751
Function CCCreateCalendarNavigator(Name, FileName,YearsRange, Calendar)
    Dim CalendarNavigator
    Set CalendarNavigator = New clsCalendarNavigator
    CalendarNavigator.Name = Name
    CalendarNavigator.FileName = FileName
    CalendarNavigator.YearsRange = YearsRange
    Set CalendarNavigator.Calendar = Calendar
    Set CCCreateCalendarNavigator = CalendarNavigator
End Function 

'End CCCreateCalendarNavigator Function

'clsCalendarNavigator Class @0-67D52788
class clsCalendarNavigator
  Public Name
  Public FileName
  Public CCSEvents
  Public Visible
  Public YearsRange

'  Public CurrentYear
'  Public CurrentMonth
'  Public CurrentDay
  Public CurrentDate
  Public CurrentProcessingDate
  Public Calendar

  Private QueryString
  Private Blocks

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set Blocks = CreateObject("Scripting.Dictionary")
    Visible = True
  End Sub

  Private Sub Class_Terminate()
    Set Blocks = Nothing
    Set CCSEvents = Nothing
  End Sub


   Private Sub SetProcessingDate(Scope, CurrentDate, Step)
      Dim Res 
      Select Case Scope
        Case calYear
          Res = DateAdd("yyyy", Step, CurrentDate)
        Case calMonth
          Res = DateAdd("m", Step, CurrentDate)
        Case cal3Month
          Res = DateAdd("m", Step, CurrentDate)
        Case calQuarter
          Res = DateAdd("q", Step, CurrentDate)
          Dim m : m = DatePart("m", Res) mod 3
          If m <>1 Then _
              Res = DateAdd("m", -IIF(m=2, 1, 2), Res)
        Case calDay
          Res = DateAdd("d", Step, CurrentDate)
        Case Else
          Res = CurrentProcessingDate 
        End Select
        CurrentProcessingDate = Res
     End Sub
  
  Private Function  CreateURL
    Dim d : d =DatePart("yyyy",CurrentProcessingDate) & "-" & CCAddZero(DatePart("m",CurrentProcessingDate), 2)
    CreateURL = FileName & "?" & CCAddParam(QueryString, Calendar.ComponentName & "Date", d)
  End Function 


  Sub SetBlockVariables(Block, Target)
    Block.Variable("URL") = CreateURL
    Block.Variable("Year") = DatePart("yyyy",CurrentProcessingDate)
    Block.Variable("Quarter") = DatePart("q", CurrentProcessingDate)
    Dim m : m =DatePart("m",CurrentProcessingDate)
    Block.Variable("Month") = m
    Block.Variable("MonthFullName") = CCSLocales.Locale.MonthNames(m-1)
    Block.Variable("MonthShortName") = CCSLocales.Locale.MonthShortNames(m-1)
    If IsEmpty(Target)Then 
       Block.Parse ccsParseAccumulate
    Else 
       Block.ParseTo ccsParseAccumulate, Target 
    End If
  End Sub 

  Sub Show(Template)
    Dim I, Block, Element, BlockName, NavigatorBlock
    Dim Years, Months, Quarters, Days 
    QueryString = CCGetQueryString("All", Array(Calendar.ComponentName & "Date", _
      Calendar.ComponentName & "Year", Calendar.ComponentName & "Month", Calendar.ComponentName & "Day", "ccsForm"))
    
    CurrentProcessingDate = Calendar.CurrentDate
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
    BlockName = "CalendarNavigator " & Name
    Set NavigatorBlock = Template.Block(BlockName)
    With NavigatorBlock
      Blocks.Add "Prev_Year", .Block("Prev_Year")
      Blocks.Add "Next_Year",.Block("Next_Year")

      If Calendar.CalendarType <> calYear Then  
        Blocks.Add "Next", .Block("Next")
        Blocks.Add "Prev", .Block("Prev")
      End If 
      
      Blocks.Add "Years", .Block("Years")
      Blocks.Add "Months", .Block("Months")
      Blocks.Add "Quarters", .Block("Quarters")
    End With

    If Visible Then
      For Each Element In Blocks
        Set Block = Blocks.Item(Element)
        If Not Block Is Nothing Then 
          If Element="Prev" Then 
            CurrentDate = Calendar.CurrentDate
            SetProcessingDate Calendar.CalendarType, CurrentDate, -1
            SetBlockVariables Block, Empty
          ElseIf Element="Next" Then 
            CurrentDate = Calendar.CurrentDate
            SetProcessingDate Calendar.CalendarType, CurrentDate, 1
            SetBlockVariables Block, Empty
          ElseIf Element="Prev_Year" Then 
            CurrentDate = Calendar.CurrentDate
            SetProcessingDate calYear, CurrentDate, -1
            SetBlockVariables Block, Empty
          ElseIf Element = "Next_Year" Then 
            CurrentDate = Calendar.CurrentDate
            SetProcessingDate calYear, CurrentDate, 1
            SetBlockVariables Block, Empty
          ElseIf Element="Years" Then  
            Block.Variable("CalendarName") = Calendar.ComponentName
            SetProcessingDate calYear, Calendar.CurrentDate, -YearsRange
            Dim y1 : y1 = DatePart("yyyy",CurrentProcessingDate)
            With Block
              Dim Current_Year : Set Current_Year = .Block("Current_Year")
              Dim Regular_Year: Set Regular_Year = .Block("Regular_Year")
            End With
            For I = 1 To YearsRange*2+1
              If y1+I-1 = Calendar.CurrentYear Then
                If Not Current_Year Is Nothing Then 
                  SetBlockVariables Current_Year, Regular_Year
                End If
              Else 
                If Not Regular_Year Is Nothing Then 
                  SetBlockVariables Regular_Year, Empty
                End If
              End If             
              SetProcessingDate calYear, CurrentProcessingDate, 1
            Next
            SetBlockVariables Block, Empty

          ElseIf Element="Months" And Calendar.CalendarType <> calYear Then  
            Block.Variable("CalendarName") = Calendar.ComponentName
            CurrentProcessingDate = DateSerial(Calendar.CurrentYear, 1, Calendar.CurrentDay)
            Set Months=Blocks("Months")
            With Months
              Dim Current_Month : Set Current_Month = .Block("Current_Month")
              Dim Regular_Month : Set Regular_Month = .Block("Regular_Month")
            End With
            For I = 1 To 12
              If I = Calendar.CurrentMonth Then
                If Not Current_Month Is Nothing Then 
                  SetBlockVariables Current_Month, Regular_Month
                End If                        
              Else 
                If Not Regular_Month Is Nothing Then 
                  SetBlockVariables Regular_Month, Empty
                End If
              End If             
              SetProcessingDate calMonth, CurrentProcessingDate, 1
            Next
            SetBlockVariables Block, Empty
          ElseIf Element="Quarters" And Calendar.CalendarType = calQuarter Then  
            Block.Variable("CalendarName") = Calendar.ComponentName
            CurrentProcessingDate = DateSerial(Calendar.CurrentYear, 1, 1)
            Set Quarters=Blocks("Quarters")
            With Quarters
              Dim Current_Quarter : Set Current_Quarter = .Block("Current_Quarter")
              Dim Regular_Quarter : Set Regular_Quarter = .Block("Regular_Quarter")
            End With
            Dim CurQuarter : CurQuarter = DatePart("q", Calendar.CurrentDate)
            For I = 1 To 4
              If I = CurQuarter Then
                If Not Current_Quarter Is Nothing Then 
                  SetBlockVariables Current_Quarter, Regular_Quarter
                End If                        
              Else 
                If Not Regular_Quarter Is Nothing Then 
                  SetBlockVariables Regular_Quarter, Empty
                End If
              End If             
              SetProcessingDate calQuarter, CurrentProcessingDate, 1
            Next
            SetBlockVariables Block, Empty
          End If
        End If
      Next

       NavigatorBlock.Variable("Action") = FileName & "?" & CCAddParam(QueryString, "ccsForm", Calendar.ComponentName)
       NavigatorBlock.Variable("CalendarName") = Calendar.ComponentName
       CurrentProcessingDate = Calendar.CurrentDate
       SetBlockVariables NavigatorBlock, Empty
    Else
      Template.Template.HideBlock(BlockName)
    End If
  End Sub

End Class

'End clsCalendarNavigator Class

'clsFileElement Class @0-CBEFE57E
Class clsFileElement
    Private mvarName
    Private mvarUploadObject
    Private mvarFileObject
    Private mvarSize
    Private mvarFileName

    Private Sub Class_Initialize()
        Set mvarUploadObject = Nothing
        Set mvarFileObject = Nothing
        mvarFileName = ""
        mvarSize = 0
    End Sub
    
    Private Sub Class_Terminate()
        Set mvarUploadObject = Nothing
        Set mvarFileObject = Nothing
    End Sub

    Public Property Set UploadObject(NewObject)
        Set mvarUploadObject = NewObject
    End Property

    Public Property Get UploadObject()
        Set UploadObject = mvarUploadObject
    End Property

    Public Property Get FileObject()
        Set FileObject = mvarFileObject
    End Property

    Public Property Let Name(ParameterName)
        If mvarUploadObject is Nothing Then Exit Property
        mvarName = ParameterName
        Set mvarFileObject = mvarUploadObject.Files(ParameterName)
        If Not mvarFileObject is Nothing Then 
            mvarFileName = mvarFileObject.FileName
            mvarSize = mvarFileObject.Size
        End If
    End Property

    Public Property Get Name()
        Name = mvarName
    End Property

    Public Property Get FileExists()
        If Not mvarFileObject is Nothing Then FileExists = (mvarFileName <> "") Else FileExists = False
    End Property

    Public Property Get Size()
        Size = mvarSize
    End Property

    Public Property Get FileName()
        FileName = mvarFileName
    End Property

    Public Function Save(NewFileName)
        Save = False
        If mvarFileObject is Nothing Then Exit Function
        mvarFileObject.SaveAs NewFileName
        Save = True
    End Function

End Class

'End clsFileElement Class

'clsFormElement Class @0-55A85CA9
Class clsFormElement
    Private mvarName
    Private mvarUploadObject
    Private mvarCount
    Private mvarValue

    Private Sub Class_Initialize()
        Set mvarUploadObject = Nothing
    End Sub
    
    Private Sub Class_Terminate()
        Set mvarUploadObject = Nothing
    End Sub

    Public Property Set UploadObject(NewObject)
        Set mvarUploadObject = NewObject
    End Property

    Public Property Get UploadObject()
        Set UploadObject = mvarUploadObject
    End Property

    Public Property Get Count()
        Count = mvarCount
    End Property

    Public Property Get Item(Index)
        Item = Empty
    End Property

    Public Property Let Name(ParameterName)
        If mvarUploadObject is Nothing Then Exit Property
        mvarValue = mvarUploadObject.Form(ParameterName).Value
        mvarCount = IIf(IsEmpty(mvarValue), 0, 1)
        mvarName = ParameterName
    End Property

    Public Property Get Name()
        Name = mvarName
    End Property

    Public Default Property Get Value() 
        Value = mvarValue
    End Property

End Class

'End clsFormElement Class

'clsUploadControl Class @0-90B4B078
Class clsUploadControl
    Private mvarUploadObject
    Private mvarFilesCount
    Private mvarFileElements
    Private mvarFormElements

    Private Sub Class_Initialize()
        mvarFilesCount = 0
        Set mvarUploadObject = Nothing
        Set mvarFileElements = CreateObject("Scripting.Dictionary")
        Set mvarFormElements = CreateObject("Scripting.Dictionary")

        On Error Resume Next

        Set mvarUploadObject = Server.CreateObject("Persits.Upload")
        mvarUploadObject.IgnoreNoPost = True

        mvarFilesCount = mvarUploadObject.Save
        If Err.Number <> 0 Then
            Response.Write "Persits uploading component Persits.Upload is not found. Please install the component or select another one."
            Response.End
        End If
        On Error Goto 0
    End Sub

    Private Sub Class_Terminate()
        mvarFileElements.RemoveAll
        Set mvarFileElements = Nothing
        mvarFormElements.RemoveAll
        Set mvarFormElements = Nothing
        Set mvarUploadObject = Nothing
    End Sub

    Public Property Get FilesCount()
      FilesCount = mvarFilesCount
    End Property

    Public Property Get Form(ParameterName)
        Dim FormElement
        If Not mvarFormElements.Exists(LCase(ParameterName)) Then
            Set FormElement = new clsFormElement
            Set FormElement.UploadObject = mvarUploadObject
            FormElement.Name = ParameterName
            mvarFormElements.Add LCase(ParameterName), FormElement
        Else
            Set FormElement = mvarFormElements.Item(LCase(ParameterName))
        End If
        Set Form = FormElement
    End Property

    Public Property Get Files(ParameterName)
      Dim FileElement
      If Not mvarFileElements.Exists(LCase(ParameterName)) Then
        Set FileElement = new clsFileElement
        Set FileElement.UploadObject = mvarUploadObject
        FileElement.Name = ParameterName
        mvarFileElements.Add LCase(ParameterName), FileElement
      Else
        Set FileElement = mvarFileElements.Item(LCase(ParameterName))
      End If
      Set Files = FileElement
    End Property
End Class

'End clsUploadControl Class

'CCCreateFileUpload Function @0-A734B792
Function CCCreateFileUpload(Name, Caption, TemporaryFolder, FileFolder, AllowedFileMasks, DisallowedFileMasks, FileSizeLimit, Required)
  Dim FileUpload
  Set FileUpload = New clsFileUpload

  With FileUpload
    .Name                = Name
    .DeleteControlName   = Name & "_Delete"
    .Caption             = Caption
    .TemporaryFolder     = TemporaryFolder & "\"
    .FileFolder          = FileFolder & "\"
    .AllowedFileMasks    = AllowedFileMasks
    .DisallowedFileMasks = DisallowedFileMasks
    .FileSizeLimit       = FileSizeLimit
    .Required            = Required            
  End With

  Set CCCreateFileUpload = FileUpload
End Function

'End CCCreateFileUpload Function

'clsFileUpload Class @0-FBF1AD42
Class clsFileUpload
  Public Name
  Public CCSEvents
  Public Visible
  Public ExternalName
  Public Errors
  Public Required
  Public TemplateBlock

  Public AllowedFileMasks
  Public DisallowedFileMasks
  Public FileSizeLimit

  Public IsUploaded
  Public FileSize
  Public fso
  Public DeleteControlName
  Public ExternalDeleteControlName

  Private mCaption

  Private VarTemporaryFolder
  Private VarFileFolder
  Private VarValue
  Private VarText
  Private StateArray
  Private IsCCSName

  Private CCSEventResult

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set fso       = CreateObject("Scripting.FileSystemObject")

    Set Errors    = New clsErrors

    ExternalName  = Empty
    Visible       = True
    IsUploaded    = False
    FileSize      = 0
    ReDim StateArray(1)
    StateArray(0) = Empty
    StateArray(1) = Empty
    IsCCSName=True
  End Sub

  Private Sub Class_Terminate()
    CCSEvents.RemoveAll
    Set CCSEvents = Nothing
    Set Errors    = Nothing
    Set fso       = Nothing
  End Sub

  Public Function Upload(CurrentRow)
  On Error Resume Next
    Dim f, FieldName, NewFileName
    If Not IsEmpty(CurrentRow) Then 
      ExternalName = Name & "_" & CStr(CurrentRow)
      ExternalDeleteControlName = Name & "_Delete_" & CStr(CurrentRow)
    End If

    If CCGetRequestParam("ccsForm",ccsGet) <> "" Then
      SetState CCGetRequestParam(IIf(Not IsEmpty(ExternalName), ExternalName, Name), ccsPost)
      Value = StateArray(0)
    End If

    If UploadedFilesCount > 0 Then
      If Not IsEmpty(objUpload.Files(IIf(Not IsEmpty(ExternalName), ExternalName, Name) & "_File")) Then _
        Set f = objUpload.Files(IIf(Not IsEmpty(ExternalName), ExternalName, Name) & "_File")
      If Not IsEmpty(f) Then
        If f.FileExists Then
          FileSize = f.Size
          NewFileName = GetValidFileName(f.FileName) & f.FileName
          CCSEventResult = CCRaiseEvent(CCSEvents, "OnRenameFile", Me)
          If TypeName(CCSEventResult)="String" And CStr(CCSEventResult)<>"" Then NewFileName = CCSEventResult
          f.Save VarTemporaryFolder & NewFileName
          If Not Err.Number = 0 Then 
            If Not CStr(Caption) = "" Then FieldName = Caption Else FieldName = Name
            Response.Write CCSLocales.GetText("CCS_UploadingTempFolderError",Array(FieldName, CStr(Err.Source) & ", " & CStr(Err.Description)))
            Response.End
          End If
          StateArray(1) = NewFileName
          If Not IsEmpty(StateArray(0)) And StateArray(1) <> StateArray(0) Then DeleteFile
          Value = NewFileName
        Else
          If IsEmpty(StateArray(0)) Then
            VarValue = ""
            FileSize = 0
          End If
        End If
      Else
        If IsEmpty(StateArray(0)) Then
          VarValue = ""
          FileSize = 0
        End If
        StateArray(1) = Empty
      End If
    End If

    If CCGetRequestParam(IIf(Not IsEmpty(ExternalDeleteControlName), ExternalDeleteControlName, DeleteControlName), ccsPost) <> "" Then DeleteFile

    If Not Err.Number = 0 Then 
      Response.Write CCSLocales.GetText("CCS_UploadingError", Array(Me.Caption, CStr(Err.Source) & ", " & CStr(Err.Description)))
      Response.End
    End If
  End Function

  Public Function GetFile(CurrentRow)
  On Error Resume Next
    Dim f, FieldName
    If Not IsEmpty(CurrentRow) Then 
      ExternalName = Name & "_" & CStr(CurrentRow)
      ExternalDeleteControlName = Name & "_Delete_" & CStr(CurrentRow)
    End If

    IsCCSName=True
    If CCGetRequestParam("ccsForm",ccsGet) <> "" Then
      SetState CCGetRequestParam(IIf(Not IsEmpty(ExternalName), ExternalName, Name), ccsPost)
      Text = StateArray(0)
      Value = Text
    End If

    If UploadedFilesCount > 0 Then
      If Not IsEmpty(objUpload.Files(IIf(Not IsEmpty(ExternalName), ExternalName, Name) & "_File")) Then _
        Set f = objUpload.Files(IIf(Not IsEmpty(ExternalName), ExternalName, Name) & "_File")
      If Not IsEmpty(f) Then
        If f.FileExists Then
          FileSize = f.Size
          Text = f.FileName
          IsCCSName = False
        Else
          If IsEmpty(StateArray(0)) Then
            Text = "" 
            FileSize = 0
          End If
        End If
      Else
        If IsEmpty(StateArray(0)) Then
          Text = "" 
          FileSize = 0
        End If
      End If
    End If

    If Not Err.Number = 0 Then 
      Response.Write CCSLocales.GetText("CCS_UploadingError", Array(Me.Caption, CStr(Err.Source) & ", " & CStr(Err.Description)))
      Response.End
    End If
  End Function

  Public Function MoveFromTempFolder
  On Error Resume Next
    Dim FieldName
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeProcessFile", Me)
    If (fso.FileExists(VarTemporaryFolder & VarValue)) Then
      fso.MoveFile VarTemporaryFolder & VarValue, VarFileFolder & VarValue
      StateArray(0) = VarValue
      StateArray(1) = VarValue
    End If
    CCSEventResult = CCRaiseEvent(CCSEvents, "AfterProcessFile", Me)
    If Not Err.Number = 0 Then
      Response.Write CCSLocales.GetText("CCS_UploadingError", Array(Me.Caption, CStr(Err.Source) & ", " & CStr(Err.Description)))
      Response.End
    End If
  End Function

  Public Function DeleteFile()
  On Error Resume Next
    Dim FieldName, FileName
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeDeleteFile", Me)
    FileName = VarValue
    If IsEmpty(FileName) Then FileName = VarText
    If (fso.FileExists(VarTemporaryFolder & FileName)) Then
      fso.DeleteFile VarTemporaryFolder & FileName, True
      VarValue = ""
    End If
    If (fso.FileExists(VarFileFolder & FileName)) Then
      fso.DeleteFile VarFileFolder & FileName, True
      VarValue = ""
    End If
    CCSEventResult = CCRaiseEvent(CCSEvents, "AfterDeleteFile", Me)
    If Not Err.Number = 0 Then 
      If Not CStr(Caption) = "" Then FieldName = Caption Else FieldName = Name
      Response.Write CCSLocales.GetText("CCS_UploadingError", Array(Me.Caption, CStr(Err.Source) & ", " & CStr(Err.Description)))
      Response.End
    End If
  End Function

  Property Let TemporaryFolder(NewText)
    Dim FieldName, regEx
    VarTemporaryFolder = NewText
    If UCase(Left(NewText, 5)) = "%TEMP" Then 
      VarTemporaryFolder = fso.GetSpecialFolder(2) & Mid(NewText, 6)
    End If
    If NewText = "" Then 
      VarTemporaryFolder = Server.MapPath(".\") & "\"
    End If
    
    Set regEx = New RegExp
    regEx.Pattern = "^[a-z]:.*"
    regEx.IgnoreCase = True
    If Not regEx.Test(VarTemporaryFolder) Then VarTemporaryFolder = Server.MapPath(VarTemporaryFolder) & "\"

    If Not fso.FolderExists(VarTemporaryFolder) Then
      Response.Write CCSLocales.GetText("CCS_TempFolderNotFound", Array(Me.Caption))
      Response.End
    End If
  End Property

  Property Get TemporaryFolder()
    TemporaryFolder = VarTemporaryFolder
  End Property

  Property Get Caption()
    If Len(Caption) > 0 Then
      Caption = mCaption
    Else
      Caption = Name
    End If
  End Property

  Property Let Caption(newValue)
    mCaption = newValue
  End Property

  Property Let FileFolder(NewText)
    Dim FieldName, regEx
    VarFileFolder = NewText
    If UCase(Left(NewText, 5)) = "%TEMP" Then 
      VarFileFolder = fso.GetSpecialFolder(2) & Mid(NewText, 6)
    End If

    Set regEx = New RegExp
    regEx.Pattern = "^[a-z]:.*"
    regEx.IgnoreCase = True
    If Not regEx.Test(VarFileFolder) Then VarFileFolder = Server.MapPath(VarFileFolder) & "\"

    If Not fso.FolderExists(VarFileFolder) Then
      Response.Write CCSLocales.GetText("CCS_FilesFolderNotFound", Array(Me.Caption))
      Response.End
    End If
  End Property

  Property Get FileFolder()
    FileFolder = CStr(VarFileFolder)
  End Property

  Property Let Value(NewValue)
  On Error Resume Next
    Dim f, FieldName

    If Not IsEmpty(NewValue) Then
      If Len(NewValue) > 0 Then

        If (fso.FileExists(VarTemporaryFolder & NewValue)) Then
          VarValue   = NewValue
          VarText  = VarValue
          StateArray(0) = NewValue
          StateArray(1) = Empty
          IsUploaded = True

          Set f = fso.GetFile(VarTemporaryFolder & NewValue)
          FileSize = f.Size
          VarText = f.Path

        ElseIf (fso.FileExists(VarFileFolder & NewValue)) Then

          VarValue   = NewValue
          VarText  = VarValue
          StateArray(0) = NewValue
          StateArray(1) = Empty
          IsUploaded = True

          Set f = fso.GetFile(VarFileFolder & NewValue)
          FileSize = f.Size

        End If

      End If
    End If    
    If Not IsEmpty(ExternalName) And NewValue = "" Then
      VarValue = ""
      VarText = ""
      StateArray(0) = Empty
      StateArray(1) = Empty
      IsUploaded = False
      FileSize = 0
    End If

    If Not Err.Number = 0 Then 
      Response.Write CCSLocales.GetText("CCS_UploadingError", Array(Me.Caption, CStr(Err.Source) & ", " & CStr(Err.Description)))
      Response.End
    End If
  End Property

  Property Get Value()
    Value = VarValue
  End Property

  Property Let Text(NewText)
    Dim f
    VarText = NewText
  End Property

  Property Get Text()
    Text = CStr(VarText)
  End Property

  Public Function GetValidFileName(FileName)
  On Error Resume Next
    Dim dta, tm, index, prefix, FieldName
    dta = Date()
    tm = time()
    index = 0

    Do
      prefix = Year(dta) & Month(dta) & Day(dta) & Hour(tm) & Minute(tm) & Second(tm) & CStr(index) & "."
      index = index + 1
    Loop While fso.FileExists(VarTemporaryFolder & prefix & FileName) Or fso.FileExists(VarFileFolder & prefix & FileName)

    GetValidFileName = prefix
    If Not Err.Number = 0 Then 
      Response.Write CCSLocales.GetText("CCS_UploadingError", Array(Me.Caption, CStr(Err.Source) & ", " & CStr(Err.Description)))
      Response.End
    End If
  End Function

  Public Function GetOriginFileName(FileName)
    Dim nPos
    nPos = InStr(FileName,".")
    If nPos > 0 Then
      GetOriginFileName = Mid(FileName,nPos+1)
    Else 
      GetOriginFileName = FileName
    End If
  End Function

  Function Validate()
    Dim FieldName,oVarText
    If Required And CStr(VarText) = "" Then
      Errors.addError(CCSLocales.GetText("CCS_RequiredFieldUpload", Me.Caption))
    End If
    oVarText = IIF(IsUploaded And IsCCSName,GetOriginFileName(VarText),VarText)
    If Not CStr(Text) = "" And DisallowedFileMasks <> "" And CCRegExpTest(oVarText, "^" & Replace(Replace(Replace(Replace(DisallowedFileMasks, ".", "\."), "?", "."), "*", ".*"), ";", "$|^") & "$", True, True) And Errors.Count = 0 Then
      Errors.addError(CCSLocales.GetText("CCS_WrongType", Me.Caption))
    End If
    If Not CStr(Text) = "" And AllowedFileMasks <> "" And AllowedFileMasks <> "*" And Not CCRegExpTest(oVarText, "^" & Replace(Replace(Replace(Replace(AllowedFileMasks,".", "\."), "?", "."), "*", ".*"), ";", "$|^") & "$", True, True) And Errors.Count = 0 Then
      Errors.addError(CCSLocales.GetText("CCS_WrongType", Me.Caption))
    End If
    If Not IsUploaded And FileSize > FileSizeLimit And Errors.Count = 0 Then
      Errors.addError(CCSLocales.GetText("CCS_LargeFile", Array(Me.Caption)))
    End If
    If Errors.Count > 0 And fso.FileExists(VarTemporaryFolder & VarText) Then DeleteFile
    Validate = CCRaiseEvent(CCSEvents, "OnValidate", Me)
  End Function

  Sub Show(Template)
    Dim TemplateBlock, UploadBlock, InfoBlock, DeleteControlBlock

    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)

    If Visible Then

      Set TemplateBlock      = Template.Block("FileUpload " & Name)
      Set UploadBlock        = TemplateBlock.Block("Upload")
      Set InfoBlock          = TemplateBlock.Block("Info")
      Set DeleteControlBlock = TemplateBlock.Block("DeleteControl")

      If Not (TemplateBlock Is Nothing) Then
          
          If IsEmpty(ExternalName) Then
            TemplateBlock.Variable("ControlName") = Name
          Else
            TemplateBlock.Variable("ControlName") = ExternalName
          End If
          TemplateBlock.Variable("State") = GetState()

          If (Not IsUploaded Or Required) And Not (UploadBlock Is Nothing) Then
            If IsEmpty(ExternalName) Then
              UploadBlock.Variable("FileControl") = Name & "_File"
            Else
              UploadBlock.Variable("FileControl") = ExternalName & "_File"
            End If
            UploadBlock.Parse ccsParseOverwrite
            InfoBlock.Visible = False
            DeleteControlBlock.Visible = False
          End If

          If IsUploaded And Not (InfoBlock Is Nothing) Then
            InfoBlock.Variable("FileName")      = GetOriginFileName(Server.HTMLEncode(VarValue))
            InfoBlock.Variable("FileSize")      = FileSize
            InfoBlock.Parse ccsParseOverwrite
            UploadBlock.Visible = Required
          End If

          If IsUploaded And Not Required And Not (DeleteControlBlock Is Nothing) Then
            If IsEmpty(ExternalDeleteControlName) Then
              DeleteControlBlock.Variable("DeleteControl") = DeleteControlName
            Else
              DeleteControlBlock.Variable("DeleteControl") = ExternalDeleteControlName
            End If
            DeleteControlBlock.Parse ccsParseOverwrite
            UploadBlock.Visible = Required
          End If

          TemplateBlock.Parse ccsParseOverwrite
      End if

    End If
  End Sub

  Function OnClick()
    OnClick = CCRaiseEvent(CCSEvents, "OnClick", Me)
  End Function

  Private Function GenerateStateKey()
    Dim dta, tm, random_number
    dta = Date()
    tm = time()
    Randomize
    random_number = Abs(Int((2147483647 + 2147483648 + 1) * Rnd - 2147483648))
    GenerateStateKey = CStr(random_number) & Day(dta) & Hour(tm) & Minute(tm) & Second(tm)
  End Function

  Private Function GetState()
    Dim ControlStateKey
    If StateArray(0) = Empty Then StateArray(0) = Value
    ControlStateKey = "FileUpload" & GenerateStateKey()
    Session(ControlStateKey) = StateArray
    GetState = ControlStateKey
  End Function

  Private Function SetState(value)  
    If IsArray(Session(value)) Then
      StateArray(0) = Session(value)(0)
      StateArray(1) = Session(value)(1)
    Else
      StateArray(0) = Empty
      StateArray(1) = Empty
    End If
  End Function

End Class

'End clsFileUpload Class

'CCCreateButton Function @0-5E520FBD
Function CCCreateButton(Name, Method)
  Dim Button
  Set Button = New clsButton
  Button.Name = Name
  If Method = ccsGet Then 
    Button.Pressed = Not IsEmpty(CCGetFromGet(Name, Empty)) Or _
                     Not IsEmpty(CCGetFromGet(Name & ".x", Empty))  Or _
                     Not IsEmpty(CCGetFromGet(Name & "_x", Empty))
  Else 
    Button.Pressed = Not IsEmpty(CCGetFromPost(Name, Empty)) Or _
                     Not IsEmpty(CCGetFromPost(Name & ".x", Empty)) Or _
                     Not IsEmpty(CCGetFromPost(Name & "_x", Empty))
  End If
  Set CCCreateButton = Button
End Function
'End CCCreateButton Function

'clsButton Class @0-3807B38E
Class clsButton
  Public Name
  Public Pressed
  Public CCSEvents
  Public Visible
  Public ExternalName

  Private CCSEventResult

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    ExternalName = Empty
    Visible = True
  End Sub

  Private Sub Class_Terminate()
    Set CCSEvents = Nothing
  End Sub

  Sub Show(Template)

    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)

    If Visible Then

      If Template.BlockExists("Button " & Name, "block") Then 
        If IsEmpty(ExternalName) Then 
          Template.Block("Button " & Name).Variable("Button_Name") = Name
        Else  
          Template.Block("Button " & Name).Variable("Button_Name") = ExternalName
        End If
        Template.Block("Button " & Name).Parse ccsParseOverwrite
      End If

    End If
  End Sub

  Function OnClick()
    OnClick = CCRaiseEvent(CCSEvents, "OnClick", Me)
  End Function
End Class

'End clsButton Class

'CCCreateDatePicker Function @0-C5EDAC24
Function CCCreateDatePicker(Name, FormName, ControlName)
  Dim DatePicker
  Set DatePicker = New clsDatePicker
  With DatePicker
    .Name = Name
    .FormName = FormName
    .ControlName = ControlName
  End With
  Set CCCreateDatePicker = DatePicker
End Function
'End CCCreateDatePicker Function

'clsDatePicker Class @0-B0B324C1
Class clsDatePicker
  Public Name
  Public ExternalName
  Public FormName
  Public ControlName
  Public ExternalControlName
  Public Visible

  Private Sub Class_Initialize()
    ExternalName = Empty
    ExternalControlName = Empty
    Visible = True
  End Sub

  Sub Show(Template)
    Dim TemplateBlock
    If Visible Then
      Set TemplateBlock = Template.Block("DatePicker " & Name)
      If Template.BlockExists("DatePicker " & Name, "block") Then 
        TemplateBlock.Variable("Name") = CStr(FormName) & "_" & CStr(Name)
        TemplateBlock.Variable("FormName") = CStr(FormName)
        If IsEmpty(ExternalControlName) Then
          TemplateBlock.Variable("DateControl") = CStr(ControlName)
        Else
          TemplateBlock.Variable("DateControl") = CStr(ExternalControlName)
        End If
        TemplateBlock.Parse ccsParseOverwrite
      End If
    End If
  End Sub

End Class
'End clsDatePicker Class

'CCCreateList Function @0-54E52904
Function CCCreateList(ControlType, Name, Caption, DataType, InitValue, DataSource)
  Dim Control
  Set Control = New clsListControl
  With Control
    .Name = Name
    .ControlType = ControlType
    .Caption = Caption
    .DataType = DataType
    .ControlTypeName = ccsControlTypes(ControlType)
    If IsArray(InitValue) Then
      .MultipleValues = InitValue
    Else
      .Text = InitValue
    End If
    If IsObject(DataSource) Then 
      Set .DataSource = DataSource
    End If
  End With
  Set CCCreateList = Control
End Function
'End CCCreateList Function

'clsListControl Class @0-9387F7A7
Class clsListControl
  Private Control
  Private DataTypeValue

  Public CCSEvents
  Public DataSource
  Public Recordset
  Public Errors
  Public Name
  Public ControlType
  Public Caption
  Public Required
  Public TemplateBlock
  Public Visible
  Public HTML
  Public MultipleValues
  Public IsMultiple
  Public ExternalName
  Public ControlTypeName
  Public PreserveVisible

  Public TextColumn
  Public BoundColumn
  Private CCSEventResult
  Private mPopulatingType

  Public IsPopulated
  Public ItemsList()
  Public KeysList()
  Public ItemsCount


  Private Sub Class_Initialize()
    Required = False
    BoundColumn = 0
    TextColumn = 1
    Visible = True
    PopulatingType = ccsStringConcats 
    HTML = False
    ExternalName = Empty
    IsPopulated = False
    ItemsCount = 0

    Set Control = New clsControl
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set Errors = New clsErrors
  End Sub

  Private Sub Class_Terminate()
    Set Control = Nothing
    Set Errors = Nothing
    Set DataSource = Nothing
    Set Recordset = Nothing
    Set CCSEvents = Nothing
  End Sub

  Public Function AddValue(NewValue)
    Dim NumberOfValues
    NumberOfValues = Ubound(MultipleValues)
    ReDim Preserve MultipleValues(NumberOfValues + 1)
    MultipleValues(NumberOfValues + 1) = NewValue
  End Function

  Public Function HasMultipleValues
    If IsArray(MultipleValues) Then
      HasMultipleValues = (Ubound(MultipleValues) > 0)
    Else
      HasMultipleValues = False
    End If
  End Function

  Property Let Value(NewValue)
    If IsMultiple Then
      If HasMultipleValues Then
        MultipleValues(1) = NewValue
      Else
        AddValue NewValue
      End If
    End If
    Control.Value = NewValue
  End Property

  Property Get Value()
    If IsMultiple And HasMultipleValues Then
      Value = MultipleValues(1)
    Else
      Value = Control.Value
    End If
  End Property

  Public Property Let PopulatingType(vType)
      mPopulatingType = vType
  End Property

  Public Property Get PopulatingType()
      PopulatingType = mPopulatingType
  End Property

  Property Let DataType(NewDataType)
    DataTypeValue = NewDataType
    Control.DataType = DataTypeValue
  End Property

  Property Get DataType()
    DataType = DataTypeValue 
  End Property

  Property Get SQLValue()
    SQLValue = Control.SQLValue
  End Property

  Function Validate()
    Dim Passed

    If Required Then
      If IsMultiple Then _
        Passed = (Ubound(MultipleValues) = 0) _
      Else _
        Passed = (CStr(Control.Value) = "") 

      If Passed Then _
        Errors.addError(CCSLocales.GetText("CCS_RequiredField", _
          Array(IIF(CStr(Caption) = "", Name, Caption))))

    End If
    Validate = CCRaiseEvent(CCSEvents, "OnValidate", Me)
  End Function


Public Sub RePopulate
    Dim cmdErrors, MaxBound, i, r, fld
    Dim RSFields
    Dim RecordsArray

    If NOT IsObject(DataSource) Then Exit Sub 

    Set cmdErrors = new clsErrors
    Set Recordset = DataSource.Exec(cmdErrors)


    Set RSFields = new clsFields
    RSFields.AddFields Array(CCCreateField("BoundColumn", BoundColumn, DataTypeValue, Empty, Recordset), _ 
      CCCreateField("TextColumn", TextColumn, ccsText, Empty, Recordset))
    If cmdErrors.Count > 0 Then
      Dim ErrorString
      If ControlType = ccsRadioButton Then
        ErrorString = "RadioButton " & Name
      ElseIf ControlType = ccsListBox Then
        ErrorString = "ListBox " & Name
      Else
        ErrorString = "CheckBoxList" & Name
      End If
      PrintDBError CCToHTML(ErrorString), "", cmdErrors.ToString()
    Else
      MaxBound = 25: i = 1
      ReDim ItemsList(MaxBound)
      ReDim KeysList (MaxBound)

      While NOT Recordset.EOF

        RecordsArray = Recordset.Recordset.GetRows(25, , Array(TextColumn,BoundColumn))

        If i >= MaxBound Then
          MaxBound = MaxBound + 25
          ReDim Preserve ItemsList(MaxBound)
          ReDim Preserve KeysList (MaxBound)
        End If

        For r = 0 To UBound(RecordsArray, 2)
         ItemsList(i) =  RSFields("TextColumn").GetParsedValue(RecordsArray(0, r),Empty)
         KeysList(i)  =  RSFields("BoundColumn").GetParsedValue(RecordsArray(1,r),Empty)
         i = i + 1
        Next
      Wend
    End If

    Recordset.Close

    IsPopulated = True
    ItemsCount = i - 1
    Set cmdErrors = Nothing

  End Sub

  Sub Show(Template)
    Dim Result, Selected, Recordset, ResultBuffer, i, j
    Dim cmdErrors
    Dim NeedShow

    If NOT IsObject(DataSource) Then Exit Sub
    If Not IsPopulated Then RePopulate()

    Set TemplateBlock = Template.Block(ControlTypeName & " " & Name)
    NeedShow = NOT (TemplateBlock Is Nothing)
    If ControlType = ccsListBox Then
      If NOT NeedShow Then _
        Set TemplateBlock = Template
    End If

    If IsEmpty(ExternalName) Then
      TemplateBlock.Variable(Name & "_Name") = Name
    Else
      TemplateBlock.Variable(Name & "_Name") = ExternalName
    End If
   
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)

    If NOT Visible Then
      If NeedShow Then TemplateBlock.Clear
      Exit Sub
    End If

    If ControlType = ccsRadioButton or ControlType = ccsCheckBoxList Then 
      TemplateBlock.Clear

      For j = 1 To ItemsCount
        Selected = ""
        If IsMultiple Then
          For i = 1 To Ubound(MultipleValues)
            If UCase(CStr(KeysList(j))) = UCase(CStr(MultipleValues(i))) Then 
              Selected = " CHECKED"
            End If
          Next
        Else
          If UCase(CStr(KeysList(j))) = UCase(CStr(Value)) Then 
            Selected = " CHECKED"
          End If
        End If
        TemplateBlock.Variable("Value") = CCSConverter.TypeToString(DataType, KeysList(j), Empty)
        TemplateBlock.Variable("Check") = Selected
        If HTML Then
          TemplateBlock.Variable("Description") = CStr(ItemsList(j))
        Else
          TemplateBlock.Variable("Description") = Server.HTMLEncode(CStr(ItemsList(j)))
        End If
        TemplateBlock.Parse True
      Next

    ElseIf ControlType = ccsListBox Then

      Set ResultBuffer = new clsStringBuffer
      Result = ""

      If mPopulatingType = ccsStringConcats Then

        For j = 1 To ItemsCount
          Selected = ""

          If IsMultiple Then
            For i = 1 To Ubound(MultipleValues)
              If UCase(CStr(KeysList(j))) = UCase(CStr(MultipleValues(i))) Then 
                Selected = " SELECTED"
                Exit For
              End If
            Next
          Else
            If UCase(CStr(KeysList(j))) = UCase(CStr(Value)) Then 
              Selected = " SELECTED"
            End If
          End If
          Result = Result & "<OPTION VALUE=""" _
            & CCSConverter.TypeToString(DataType, KeysList(j), Empty) _
            & """" & Selected & ">" & Server.HTMLEncode(ItemsList(j)) & "</OPTION>" & vbNewLine
        Next
      Else
        For j = 1 To ItemsCount
          Selected = ""
          If IsMultiple Then
            For i = 1 To Ubound(MultipleValues)
              If UCase(CStr(KeysList(j))) = UCase(CStr(MultipleValues(i))) Then 
                Selected = " SELECTED"
                Exit For
              End If
            Next
          Else
            If UCase(CStr(KeysList(j))) = UCase(CStr(Value)) Then 
              Selected = " SELECTED"
            End If
          End If
          ResultBuffer.Append "<OPTION VALUE=""" _
            & CCSConverter.TypeToString(DataType, KeysList(j), Empty) _
            & """" & Selected & ">" & Server.HTMLEncode(ItemsList(j)) & "</OPTION>" & vbNewLine
        Next
        Result = ResultBuffer.ToString
      End If
      TemplateBlock.Variable(Name & "_Options") = Result
      If NeedShow Then TemplateBlock.Show
    End If
    Set TemplateBlock = Nothing
  End Sub

  Property Get Text()
    Text = Control.Text
  End Property

  Property Let Text(NewText)
    Control.Text = NewText
  End Property

  Property Get SQLText()
    SQLText = Control.SQLText
  End Property

  Property Let SQLText(NewSQLText)
    Control.SQLText = NewSQLText
  End Property

End Class
'End clsListControl Class

'clsErrors Class @0-D9197FC1
Class clsErrors
  Private ErrorsCount
  Private Errors
  Public ErrorDelimiter

  Private Sub Class_Initialize()
    Clear
    ErrorDelimiter = "<br>"
  End Sub

  Sub AddError(Description)
    If NOT(CStr(Description) = "") Then
      ReDim Preserve Errors(ErrorsCount)
      Errors(ErrorsCount) = Description
      ErrorsCount = ErrorsCount + 1
    End If
  End Sub

  Sub AddErrors(objErrors)
    Dim I
    For I = 0 To objErrors.Count - 1
      AddError(objErrors.ErrorByNumber(I))
    Next
  End Sub

  Sub Clear()
    ErrorsCount = 0
    ReDim Errors(1)
  End Sub

  Property Get Count()
    Count = ErrorsCount
  End Property

  Property Get ErrorByNumber(ErrorNumber)
    If ErrorNumber > ErrorsCount OR ErrorNumber < 0 Then 
      Err.Raise 4001, "Error class, ErrorByNumber function. Parameter out of range."
    End If
    ErrorByNumber = Errors(ErrorNumber)
  End Property

  Property Get ToString()
    If ErrorsCount > 0 Then
      ToString = Join(Errors, ErrorDelimiter)
    Else
      ToString = ""
    End If
  End Property

End Class
'End clsErrors Class

'clsEventCaller Class @0-164C2E9C
Class clsEventCaller
  Public EventRef
  Function Invoke(EventAgrgs)
    Invoke = Me.EventRef(EventAgrgs)
  End Function
End Class
'End clsEventCaller Class

'CCCreateDataSource Function @0-07E5CF26
Function CCCreateDataSource(DataSourceType, Connection, CommandSource)
  Dim Cmd
  
  Set Cmd = New clsCommand
  If DataSourceType <> dsListOfValues Then
    Set Cmd.Connection = Connection
    Set Cmd.WhereParameters.Connection = Connection
  End If
  Cmd.CommandType = DataSourceType
  Cmd.CommandOperation = cmdOpen
  Cmd.ActivePage = -1

  Select Case DataSourceType
    Case dsTable
      Cmd.SQL = CommandSource(0)
      Cmd.Where = CommandSource(1)
      Cmd.OrderBy = CommandSource(2)
    Case dsSQL
      Cmd.SQL = CommandSource
    Case dsProcedure
      Set Cmd.SQL = CommandSource
    Case dsListOfValues
      Cmd.LOV = CommandSource
  End Select
  Set CCCreateDataSource = Cmd
End Function
'End CCCreateDataSource Function

'clsEmptyDataSource Class @0-31A700B1
Class clsEmptyDataSource

  Public Errors
  Public CCSEvents

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set Errors = New clsErrors
  End Sub

  Private Sub Class_Terminate()
    Set CCSEvents = Nothing
    Set Errors = Nothing
  End Sub

  Function Open(Cmd)
    Set Open = Me
  End Function

  Property Get EOF()
    EOF = True
  End Property

  Property Get BOF()
    BOF = True
  End Property

  Property Get State()
    State = adStateClosed
  End Property

  Property Get Fields(Name)
    Fields = Empty
  End Property

  Property Get EditMode(ReadAllowed)
    EditMode = False
  End Property

End Class

'End clsEmptyDataSource Class

'clsDataSource Class @0-BF546957
Class clsDataSource

  Public DataSourceType
  Public DataSource
  Public Errors, Connection, Parameters, CCSEvents
  Public Recordset
  Public PageSize
  Public Command

  Private mRecordCount
  Public Order
  Public StaticOrder
  Public HasOutputParameters

  Private objFields
  Private AbsolutePage
  Private builtSQL
  Private Opened
  Private MemoFields

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set MemoFields = CreateObject("Scripting.Dictionary") 
    Set Errors = New clsErrors
    Set Parameters = New clsSQLParameters
    Set Parameters.DataSource = Me
    AbsolutePage = 0
    RecordCount = -1
    Opened = False
  End Sub

  Sub Close()
    If Recordset.State = adStateOpen Then 
      Recordset.Close
    End If
    Opened = False
    Set Recordset = Nothing
  End Sub

  Function GetData(Pages)
    If Pages = 0 Then
      Set GetData = Recordset
    End If
  End Function

  Property Get Fields(Name)
    If IsNumeric(Name) Then
      Fields = CCGetValue(Recordset, CInt(Name))
    ElseIf IsObject(objFields) Then
      If Not objFields is Nothing Then
        If objFields.Exists(Name) Then
          If MemoFields.Exists(objFields(Name).DBFieldName) Then
            Fields = MemoFields(objFields(Name).DBFieldName)
          Else
            Fields = objFields(Name).Value
            If objFields(Name).DataType = ccsMemo Then
              MemoFields.Add objFields(Name).DBFieldName, Fields
            End If
          End If
        Else
          Fields = CCGetValue(Recordset, Name)
        End If
      Else
        Fields = CCGetValue(Recordset, Name)
      End If
    Else
      Fields = CCGetValue(Recordset, Name)
    End If
  End Property


  Property Set FieldsCollection(NewFieldsCollection)
    Set objFields = NewFieldsCollection
    If Not objFields Is Nothing Then
      objFields.InitEnum
      While Not objFields.EndOfEnum
        Set objFields.NextItem.DataSource = Me
      Wend
    End If
  End Property
  
  Property Get EOF()
    EOF = Recordset.EOF
  End Property

  Property Get BOF()
    BOF = Recordset.BOF
  End Property


  Property Get State()
    If IsObject(Recordset) Then
      State = Recordset.State
    Else
      State = False
    End If
  End Property

  Sub MoveNext()
    Recordset.MoveNext
    MemoFields.RemoveAll
  End Sub

  Sub MoveFirst()
    Recordset.MoveFirst
    MemoFields.RemoveAll
  End Sub

  Function GetOrder(DefaultSorting, Sorter, Direction, MapArray)
    Dim OrderValue, I, ActiveSorter

    If NOT IsEmpty(Sorter) Then
      ' Select sorted column
      I = 0
      Do While I <= UBound(MapArray)
        If MapArray(I)(0) = Sorter Then
          ActiveSorter = I
          Exit Do
        End If
        I = I + 1
      Loop
      If NOT IsEmpty(ActiveSorter) Then
        If NOT IsEmpty(Direction) AND (Direction = "ASC" OR Direction = "DESC") Then
          If Direction = "ASC" Then
            OrderValue = MapArray(ActiveSorter)(1)
          ElseIf Direction = "DESC" Then
            OrderValue = MapArray(ActiveSorter)(2)
          End If
          If OrderValue = "" Then 
            OrderValue = MapArray(ActiveSorter)(1) & " DESC"
          End If
        Else
          OrderValue = MapArray(ActiveSorter)(1)
        End If
      End If
    End If
    If Len(OrderValue) > 0 Then
      Order = OrderValue
    Else
      Order = DefaultSorting
    End If
    If(Len(StaticOrder)>0) Then
      If Len(Order)>0 Then Order = ", "+Order
      Order = StaticOrder + Order
    End If
    GetOrder = Order
  End Function

  Public Property Let RecordCount(vData)
    mRecordCount = vData
  End Property

  Public Property Get RecordCount()
    If mRecordCount < 0 Then 
      mRecordCount = Command.ExecuteCount
    End If
    RecordCount = mRecordCount
  End Property

  Function MoveToPage(Page)
    Dim PageCounter
    Dim RecordCounter

    If Recordset.State = adStateOpen Then
      PageCounter = 1
      RecordCounter = 1
      While NOT Recordset.EOF AND PageCounter < Page
        If RecordCounter MOD Command.PageSize = 0 Then 
          PageCounter = PageCounter + 1
        End If

        RecordCounter = RecordCounter + 1
        Recordset.MoveNext
      Wend
    End If
    Command.ActivePage = PageCounter
  End Function

  Function PageCount()
    Dim Result
    If Command.PageSize > 0 Then
      Result = RecordCount \ Command.PageSize
      If (RecordCount MOD Command.PageSize) > 0 Then 
        Result = Result + 1
      End If
    Else
      Result = 1
    End If
    PageCount = Result
  End Function

  Public Function EditMode(ReadAllowed)
     If Me.State = adStateClosed And HasOutputParameters Then 
          EditMode = True
     ElseIf Me.State = adStateOpen Then 
       If Me.State = adStateOpen Then 
             EditMode = NOT Recordset.EOF And ReadAllowed
       Else 
            EditMode = False
       End If
     End If
  End Function

  Public Function CanPopulate()
    If Me.State =adStateClosed And HasOutputParameters Then 
       CanPopulate = True
    Else 
       CanPopulate = Not Me.Recordset.EOF 
    End If
  End Function


  Private Sub Class_Terminate()
    Set Command = Nothing
    Set Errors = Nothing
    Set Parameters = Nothing
    Set CCSEvents = Nothing
  End Sub

End Class

'End clsDataSource Class

'clsCommand Class @0-1C5FC536
Class clsCommand

  Private mCommandType
  Private mCommandOperation
  Private mPrepared

  Private mSQL
  Private mCountSQL
  Private mWhere
  Private mOrderBy
  Private mLOV
  Private mSP
  
  Private mPageSize
  Private mActivePage

  Private mFirstPartSQL
  Private mSecondPartSQL
  
  Public Errors, Connection, CCSEvents
  Public WhereParameters, Parameters
  Public CommandParameters
  Public Options
  Private IsNeedMoveToPage
  Public RecordsetNumber

  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set WhereParameters = New clsSQLParameters
    Set WhereParameters.ParameterSources = CreateObject("Scripting.Dictionary")
    Set Parameters = New clsSQLParameters
    Set Parameters.ParameterSources = CreateObject("Scripting.Dictionary")
    Set Options = CreateObject("Scripting.Dictionary")
    ActivePage = 0
    Prepared = False
    IsNeedMoveToPage=True
    mCountSQL = Empty
    RecordsetNumber = 0 
  End Sub

  Public Function Exec(Err)
    Set Errors = Err
    Select Case CommandOperation
      Case cmdOpen
        Set Exec = DoOpen
      Case cmdExec
        DoExec
    End Select
  End Function

  Private Function OpenRecordset(sSQL,isCountSQL)
    Dim Command
    Dim Recordset

    If Not isCountSQL And Options.Count>0 Then
      Dim Page : Page=IIF(mActivePage>0,mActivePage,1)
      Dim Size : Size=IIF(mPageSize>0,mPageSize,1)
      If Options.Exists("TOP") Then 
           sSQL=Replace(sSQL,"{SqlParam_endRecord}", Page * Size, 1, 1)
      End If
      If Options.Exists("LIMIT MYSQL") Then
           sSQL = sSQL & " LIMIT " & (Page - 1) * Size  & " , " & Size
           IsNeedMoveToPage=False
      End If
      If Options.Exists("LIMIT POSTGRES") Then
           sSQL = sSQL & " LIMIT " & Size  & " OFFSET " & (Page - 1) * Size
           IsNeedMoveToPage=False
      End If
   End If


   Set Command = CreateObject("ADODB.Command")
    Command.CommandType = adCmdText
    Command.CommandText = sSQL
    Set Command.ActiveConnection = Connection.Connection
    Set Recordset = Connection.Execute(Command)
    
    If Connection.Errors.Count > 0 Then 
      Errors.AddErrors Connection.Errors
    End If

    Set OpenRecordset = Recordset
    Set Command = Nothing
  End Function

  
  Private Function ParseParams(sSQL, Params)
    Dim I
    Dim NewSQL
    Dim ParamKeys
    Dim ParamItems
    
    NewSQL = sSQL
    If CommandType = dsSQL Then
      If Not Params is Nothing Then
        ParamItems = Params.ParametersList.Items
        ParamKeys = Params.ParametersList.Keys
        For I = 0 To UBound(ParamItems)
         NewSQL = Replace(NewSQL, "{" & ParamKeys(I) & "}", ParamItems(I).SQLText)
        Next
      End If
    End If
    ParseParams = NewSQL
  End Function

  Private Function DoOpen()
    Dim Command
    Dim builtSQL
    Dim DataSource
    Dim CountRecordset
    Dim ResultRecordset
    Dim CCSEventResult
    Dim ParameterValue
    Dim Parameter

    Set DataSource = new clsDataSource
    If IsObject(Connection) Then _
      Set DataSource.Connection = Connection
    Set DataSource.Command = Me
    
    Select Case CommandType
      Case dsTable, dsSQL
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildSelect", Me)
        If  InStr(SQL, "{SQL_Where}") > 0 Or InStr(SQL,"{SQL_OrderBy}") > 0 Then
          SQL = Replace(SQL, "{SQL_Where}", IIf(Len(Where) > 0, " WHERE " & Where, ""))
          If InStr(SQL,"{SQL_OrderBy}") > 0 Then 
             SQL = Replace(SQL, "{SQL_OrderBy}", IIf(Len(OrderBy) > 0, " ORDER BY " & OrderBy, ""))
          Else 
             SQL = SQL & IIf(Len(OrderBy) > 0, " ORDER BY " & OrderBy, "")
          End If
          builtSQL = ParseParams(SQL, WhereParameters)
        Else 
          builtSQL = ParseParams(SQL & IIf(Len(Where) > 0, " WHERE " & Where, "") & IIf(Len(OrderBy) > 0, " ORDER BY " & OrderBy, ""), WhereParameters)
        End If
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteSelect", Me)
        Set DataSource.Recordset = OpenRecordset(builtSQL,False)
        If ActivePage > 0 And IsNeedMoveToPage Then 
          DataSource.MoveToPage ActivePage
        End If
        If Not IsEmpty(mCountSQL) And Len(CountSQL) = 0 Then
          If DataSource.Recordset.State = adStateOpen Then
            Dim Counter : Counter = 0
            While NOT DataSource.Recordset.EOF AND Counter < mPageSize+1
               Counter = Counter + 1
               DataSource.Recordset.MoveNext
            Wend
            DataSource.RecordCount = IIF(ActivePage>0,(ActivePage - 1) * mPageSize, 0) + Counter
          End If 
          If ActivePage > 0 Then
            DataSource.Recordset.MoveFirst
            DataSource.MoveToPage ActivePage
          End If
    End If 


        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteSelect", Me)
        If Errors.Count > 0 Then 
          DataSource.Errors.AddErrors Errors
        End If
        Set DoOpen = DataSource

     Case dsProcedure
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildSelect", Me)
        Set Command = CreateSP()
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteSelect", Me)
        Set DataSource.Recordset = Connection.Execute(Command)

        If IsArray(CommandParameters) Then
          For I = 0 To UBound(CommandParameters)
            If IsArray(CommandParameters(I)) Then
              If CommandParameters(I)(3)=adParamInputOutput Or _
                CommandParameters(I)(3)=adParamOutput Or _
                CommandParameters(I)(3)=adParamReturnValue Then 
                  DataSource.HasOutPutParameters = True
                Exit For
              End If
            End If
          Next

          If Connection.Database = "MSSQLServer" Then
            If DataSource.HasOutPutParameters Then 
              While (Not DataSource.Recordset is Nothing)
                Set DataSource.Recordset=DataSource.Recordset.NextRecordset
              Wend
            End If
          End If

          For I = 0 To UBound(CommandParameters)
            If IsArray(CommandParameters(I)) Then
              If CommandParameters(I)(3)=adParamInputOutput Or _
                CommandParameters(I)(3)=adParamOutput Or _
                CommandParameters(I)(3)=adParamReturnValue Then 
                  Parameters.ParameterSources(CommandParameters(I)(1))=Command.Parameters(CommandParameters(I)(0))
                DataSource.HasOutPutParameters = True
              End If
            End If
          Next
            
          If Connection.Database = "MSSQLServer" Then
            If DataSource.HasOutPutParameters Then
              Set DataSource.Recordset = Connection.Execute(Command)
            End If
          End If
        End If

        If ActivePage > 0 Then
          DataSource.MoveToPage ActivePage
        End If

        If DataSource.Recordset.State > 0 Then 
          Dim TempNumber : TempNumber = 0
          Do While TempNumber <> RecordsetNumber
            Set DataSource.Recordset = DataSource.Recordset.NextRecordset
            If  DataSource.Recordset is Nothing Then 
              DataSource.Errors.AddError("wrong RecordsetNumber")
              Exit Do
            End If 
            TempNumber = TempNumber + 1
          Loop
        End If

        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteSelect", Me)
        If Connection.Errors.Count > 0 Then 
          DataSource.Errors.AddErrors Connection.Errors
        End If

        Set Command = Nothing
        Set DoOpen = DataSource

      Case dsListOfValues
        Dim I
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildSelect", Me)
        Set DataSource.Recordset = CreateObject("ADODB.Recordset")
        DataSource.Recordset.Fields.Append "bound", adBSTR, 256, adFldCacheDeferred + adFldUpdatable
        DataSource.Recordset.Fields.Append "text", adBSTR, 256, adFldCacheDeferred + adFldUpdatable
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteSelect", Me)
        DataSource.Recordset.Open
        For I = 0 To UBound(mLOV(0))
          DataSource.Recordset.AddNew
          DataSource.Recordset.Fields("bound").Value = mLOV(0)(I)
          DataSource.Recordset.Fields("text").Value = mLOV(1)(I)
        Next
        DataSource.Recordset.Update
        DataSource.Recordset.MoveFirst
        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteSelect", Me)
        Set DoOpen = DataSource
    End Select
  End Function

  Public Function ExecuteCount()
    Dim Result: Result = 0
    Dim builtSQL: builtSQL = ""
    Dim CountRecordset

    If Len(CountSQL) > 0 Then
      If  InStr(CountSQL, "{SQL_Where}") > 0 Or InStr(CountSQL,"{SQL_OrderBy}") > 0 Then
        builtSQL = ParseParams(Replace(Replace(CountSQL, "{SQL_Where}", IIf(Len(Where) > 0, " WHERE " & Where, "")),"{SQL_OrderBy}",""), WhereParameters)
      Else 
        builtSQL = ParseParams(CountSQL & IIf(Len(Where) > 0, " WHERE " & Where, ""), WhereParameters)
      End If
      Set CountRecordset = OpenRecordset(builtSQL,True)
      If CountRecordset.State = adStateOpen Then 
        Result = CLng(CountRecordset.Fields(0).Value)
      End If
      Set CountRecordset = Nothing
    End If
    ExecuteCount = Result
  End Function

  Private Function CreateSP()
    Dim Command, I, ParameterValue, Parameter, Sources

    Set Command = Server.CreateObject("ADODB.Command")
    Set Command.ActiveConnection = Connection.Connection
    Command.CommandType = adCmdStoredProc
    Command.CommandText = mSP
    If IsArray(CommandParameters) Then
      Set Sources = Parameters.ParameterSources
      For I = 0 To UBound(CommandParameters)
        ParameterValue = Sources(CommandParameters(I)(1))
        Dim CCSParType : CCSParType = GetCCSType(CommandParameters(I)(2))
        If IsEmpty(ParameterValue) OR (ParameterValue="" AND (CCSParType = ccsInteger OR CCSParType = ccsFloat OR CCSParType = ccsSingle OR CCSParType = ccsDate)) Then
          ParameterValue = CommandParameters(I)(7)
        End If
        If IsEmpty(ParameterValue) Then 
          ParameterValue = Null
        ElseIf VarType(ParameterValue) = vbString Then 
          ParameterValue = CCSConverter.StringToType(CCSParType, ParameterValue, CommandParameters(I)(8))
        End If
        Set Parameter = Command.CreateParameter(CommandParameters(I)(0), CommandParameters(I)(2), CommandParameters(I)(3), CommandParameters(I)(4), ParameterValue)
        If Parameter.Type = adNumeric Then
          Parameter.NumericScale = CommandParameters(I)(5)
          Parameter.Precision = CommandParameters(I)(6)
        End If
        Command.Parameters.Append Parameter
      Next
      Set Sources = Nothing
    End If
    Set CreateSP = Command
  End Function

  Private Sub DoExec
    Dim Command, I
    Dim builtSQL
    Dim ParameterValue, ParameterLength

    If CommandType = dsProcedure Then
      Set Command = CreateSP()
    Else
      Set Command = CreateObject("ADODB.Command")
      Command.CommandType = adCmdText

      builtSQL = SQL
      If CommandType = dsSQL Then
        builtSQL = ParseParams(builtSQL, WhereParameters)
        builtSQL = ParseParams(builtSQL, Parameters)
      Else
        If IsArray(CommandParameters) Then
          For I = 0 To UBound(CommandParameters)
	    If IsArray(CommandParameters(I)) Then
              If IsEmpty(CommandParameters(I)(4)) Then 
                ParameterValue = Null 
                ParameterLength = 1
              Else 
                ParameterValue = CommandParameters(I)(4)
                ParameterLength = CommandParameters(I)(3)
              End If
              Command.Parameters.Append Command.CreateParameter(CommandParameters(I)(0), CommandParameters(I)(1), CommandParameters(I)(2), ParameterLength, ParameterValue)
            End If
          Next
        End If
      End If

      Command.CommandText = builtSQL
      Command.Prepared = Prepared

      Set Command.ActiveConnection = Connection.Connection
    End If
    Connection.Execute(Command)
  
    If CommandType = dsProcedure And Connection.Errors.Count = 0 Then 
      If IsArray(CommandParameters) Then
        For I = 0 To UBound(CommandParameters)
          If CommandParameters(I)(3)=adParamInputOutput Or _
            CommandParameters(I)(3)=adParamOutput Or _
            CommandParameters(I)(3)=adParamReturnValue Then 
              Parameters.ParameterSources(CommandParameters(I)(1))=Command.Parameters(CommandParameters(I)(0))
'                DataSource.HasOutPutParameters = True
          End If
        Next
      End If
    End If

    If Connection.Errors.Count > 0 Then 
      Errors.AddErrors Connection.Errors
    End If

    Set Command = Nothing
  End Sub

  Public Sub AddSQLStrings(FirstPart, SecondPart)
    If Not IsEmpty(mFirstPartSQL) Then mFirstPartSQL = mFirstPartSQL & ", "
    mFirstPartSQL = mFirstPartSQL & FirstPart
    If Not IsEmpty(mSecondPartSQL) Then mSecondPartSQL= mSecondPartSQL & ", "
    mSecondPartSQL = mSecondPartSQL & SecondPart
  End Sub

  Public Function PrepareSQL(What, Table, TableWhere)
    If  mFirstPartSQL="" Then 
	  PrepareSQL = False 
    Else 
     Select Case What
       Case "Insert"
         mSQL = "INSERT INTO " & Table & "("  & mFirstPartSQL &  ") VALUES (" & mSecondPartSQL &  ")"
       Case "Update"
	 mSQL = "UPDATE " & Table & " SET " & mFirstPartSQL & IIf(Len(TableWhere) > 0, " WHERE " & TableWhere, "")
       End Select 
       PrepareSQL = True
    End If
    mFirstPartSQL = Empty
    mSecondPartSQL = Empty
  End Function

  Public Property Let ActivePage(vData)
      mActivePage = vData
  End Property

  Public Property Get ActivePage()
      ActivePage = mActivePage
  End Property

  Public Property Let PageSize(vData)
      mPageSize = vData
  End Property

  Public Property Get PageSize()
      PageSize = mPageSize
  End Property

  Public Property Let CommandOperation(vData)
      mCommandOperation = vData
  End Property

  Public Property Get CommandOperation()
      CommandOperation = mCommandOperation
  End Property

  Public Property Let LOV(vData)
      mLOV = vData
  End Property

  Public Property Get LOV()
      LOV = mLOV
  End Property

  Public Property Let SP(vData)
      mSP = vData
  End Property

  Public Property Get SP()
      SP = mSP
  End Property

  Public Property Let CountSQL(vData)
      mCountSQL = vData
  End Property

  Public Property Get CountSQL()
      CountSQL = mCountSQL
  End Property

  Public Property Let SQL(vData)
      mSQL = vData
  End Property

  Public Property Get SQL()
      SQL = mSQL
  End Property

  Public Property Let Prepared(vData)
      mPrepared = vData
  End Property

  Public Property Get Prepared()
      Prepared = mPrepared
  End Property

  Public Property Let CommandType(vData)
      mCommandType = vData
  End Property

  Public Property Get CommandType()
      CommandType = mCommandType
  End Property

  Public Property Let OrderBy(vData)
      mOrderBy = vData
  End Property

  Public Property Get OrderBy()
      OrderBy = mOrderBy
  End Property

  Public Property Let Order(vData)
      mOrderBy = vData
  End Property

  Public Property Get Order()
      Order = mOrderBy
  End Property

  Public Property Let Where(vData)
      mWhere = vData
  End Property

  Public Property Get Where()
      Where = mWhere
  End Property

  Public Sub Class_Terminate()
    Set Options = Nothing
    Set CCSEvents = Nothing
  End Sub

End Class

'End clsCommand Class

'clsConverter Class @0-916B5E6D
Class clsConverter
  Private mDateFormat
  Private mBooleanFormat
  Private mIntegerFormat
  Private mFloatFormat
  Private mSingleFormat
  Private mParseError

  Private Sub Class_Initialize()
    mParseError = False
  End Sub

  Property Let DateFormat(newDateFormat)
    mDateFormat = newDateFormat
  End Property

  Property Get DateFormat()
    DateFormat = mDateFormat
  End Property

  Property Let BooleanFormat(newFormat)
    mBooleanFormat = newFormat
  End Property

  Property Get BooleanFormat()
    BooleanFormat = mBooleanFormat
  End Property

  Property Let IntegerFormat(newFormat)
    mIntegerFormat = newFormat
  End Property

  Property Get IntegerFormat()
    IntegerFormat = mIntegerFormat
  End Property

  Property Let FloatFormat(newFormat)
    mFloatFormat = newFormat
  End Property

  Property Get FloatFormat()
    FloatFormat = mFloatFormat
  End Property

  Property Let SingleFormat(newFormat)
    mSingleFormat = newFormat
  End Property

  Property Get SingleFormat()
    SingleFormat = mSingleFormat
  End Property

  Property Get ParseError()
    ParseError = mParseError
  End Property

  Public Function VBSConvert(DataType, Value)
    Dim Result

    mParseError = False
    Result = Empty

    If IsEmpty(Value) Then
      VBSConvert = Result
      Exit Function
    End If
    
    On Error Resume Next
    Select Case DataType
      Case ccsDate
        Result = CDate(Value)
      Case ccsBoolean
        Result = CBool(Value)
      Case ccsInteger
        Result = CLng(Value)
      Case ccsFloat
        Result = CDbl(Value)
      Case ccsSingle
        Result = CSng(Value)
      Case ccsText, ccsMemo
        Result = CStr(Value)
    End Select
    If Err.Number <> 0 Then _
      mParseError = True
    On Error Goto 0
    VBSConvert = Result
  End Function

  Public Function StringToType(DataType, Value, Format)
    Dim CurrentFormat
    Dim Result

    mParseError = False
    Result = Empty
    If IsEmpty(Value) Then
      StringToType = Result
      Exit Function
    End If
    If IsEmpty(Format) Then
      Select Case DataType
        Case ccsDate
          CurrentFormat = mDateFormat    
        Case ccsBoolean
          If IsEmpty(mBooleanFormat) Then 
             CurrentFormat = CCSLocales.Locale.BooleanFormat
          Else
             CurrentFormat = mBooleanFormat
          End If
        Case ccsInteger
          CurrentFormat = mIntegerFormat
        Case ccsFloat
          CurrentFormat = mFloatFormat
        Case ccsSingle
          CurrentFormat = mSingleFormat
      End Select
    Else
      If CCSLocales.Locale.OverrideDateFormats Then 
        Select Case Format(0)
          Case "LongDate" CurrentFormat = CCSLocales.Locale.LongDate
          Case "LongTime" CurrentFormat = CCSLocales.Locale.LongTime
          Case "ShortDate" CurrentFormat = CCSLocales.Locale.ShortDate
          Case "ShortTime" CurrentFormat = CCSLocales.Locale.ShortTime
          Case "GeneralDate" CurrentFormat = CCSLocales.Locale.GeneralDate
          Case Else  CurrentFormat = Format
        End Select
      Else 
         CurrentFormat = Format
      End If 
    End If

    On Error Resume Next
    Select Case DataType
      Case ccsDate
        Result = CCParseDate(Value, CurrentFormat)
      Case ccsBoolean
        Result = CCParseBoolean(Value, CurrentFormat)
      Case ccsInteger
        Result = CCParseInteger(Value, CurrentFormat)
      Case ccsFloat
        Result = CCParseFloat(Value, CurrentFormat)
      Case ccsSingle
        Result = CCParseSingle(Value, CurrentFormat)
      Case ccsText, ccsMemo
        Result = CStr(Value)
    End Select
    If Err.Number <> 0 Then _
      mParseError = True
    On Error Goto 0
    StringToType = Result
  End Function

  Public Function TypeToString(DataType, Value, Format)
    Dim CurrentFormat
    Dim Result
    Dim VarDataType
    Dim CurrentDataType

    VarDataType = VarType(Value)
    Select Case VarDataType
      Case vbInteger, vbLong, vbByte:
        CurrentDataType = ccsInteger
      Case vbDouble, vbCurrency, vbDecimal:
        CurrentDataType = ccsFloat
      Case vbSingle:
        CurrentDataType = ccsSingle
      Case vbDate:
        CurrentDataType = ccsDate
      Case vbBoolean:
        CurrentDataType = ccsBoolean
      Case vbString:
        CurrentDataType = ccsText
      Case vbArray, vbDataObject, vbVariant, vbError, vbObject, vbNull:
        Err.Raise 1057, "Type mismatch"
    End Select
    If DataType = ccsMemo AND CurrentDataType = ccsText Then _
       CurrentDataType = ccsMemo
    If (VarDataType <> vbEmpty) AND (CurrentDataType <> DataType) Then _
        Err.Raise 1057, "Type mismatch"

    If IsEmpty(Format) Then
      Select Case DataType
        Case ccsDate
              CurrentFormat = mDateFormat
        Case ccsBoolean
          If IsEmpty(mBooleanFormat) Then 
             CurrentFormat = CCSLocales.Locale.BooleanFormat
          Else
             CurrentFormat = mBooleanFormat
          End If
        Case ccsInteger
          CurrentFormat = mIntegerFormat
        Case ccsFloat
          CurrentFormat = mFloatFormat
        Case ccsSingle
          CurrentFormat = mSingleFormat
      End Select
    ElseIf DataType=ccsDate And UBound(Format)=0 And CCSLocales.Locale.OverrideDateFormats Then 
        Select Case Format(0)
          Case "LongDate" CurrentFormat = CCSLocales.Locale.LongDate
          Case "LongTime" CurrentFormat = CCSLocales.Locale.LongTime
          Case "ShortDate" CurrentFormat = CCSLocales.Locale.ShortDate
          Case "ShortTime" CurrentFormat = CCSLocales.Locale.ShortTime
          Case "GeneralDate" CurrentFormat = CCSLocales.Locale.GeneralDate
          Case Else  CurrentFormat = Format
        End Select
    Else 
         CurrentFormat = Format
    End If 
    Select Case DataType
      Case ccsDate
        Result = CCFormatDate(Value, CurrentFormat)
      Case ccsBoolean
        Result = CCFormatBoolean(Value, CurrentFormat)
      Case ccsInteger
        Result = CCFormatNumber(Value, CurrentFormat)
      Case ccsFloat
        Result = CCFormatNumber(Value, CurrentFormat)
      Case ccsSingle
        Result = CCFormatNumber(Value, CurrentFormat)
      Case ccsText, ccsMemo
        Result = CStr(Value)
    End Select
    TypeToString = Result
  End Function
End Class
'End clsConverter Class

'clsLocales Class @0-19F6B7A8
Const CCS_I18N_RequestParameterName = 0
Const CCS_I18N_CookieName = 1
Const CCS_I18N_SessionName = 2
Const CCS_I18N_LanguageSessionName = 3 
Const CCS_I18N_HttpHeaderName = 4

Class clsLocales
  Public Locales
  Public AppPrefix
  Public Locale
  Public UseStaticTranslation   

  Private mDefaultLanguage
  Private mLanguage
  Private mCountry
  Private mCache
  Private mPathRes   
  Private FilePath
  Private Ext
  Private Keys,Vals
  Private CachePrefix
  Private IsFallback
  Private MainLanguage  
  Private Prepared  
  Private DateLastModified 
  Private CachePrefixLocales   
  Private CanUseCache   

  Private Sub Class_Initialize()
    Dim FSO
    Ext = ".txt"
    mCache = True
    IsFallback = True
    Prepared = ""
    UseStaticTranslation = False
    'DeleteTranslations
    Set Locales = Server.CreateObject("Scripting.Dictionary")
    Set Locale = Nothing
    CanUseCache = True
  End Sub

  Private Sub Class_Terminate()
    Set Locales = Nothing
  End Sub

  Property Get DefaultLanguage()
    DefaultLanguage = mDefaultLanguage
  End Property

  Property Let DefaultLanguage(NewLanguage)
    mDefaultLanguage = NewLanguage
  End Property

  Property Get PathRes()
    PathRes = mPathRes
  End Property

  Property Let PathRes(NewPathRes)
    If Len(NewPathRes) = 0 Or Right(NewPathRes,1) <> "\" Then _
      NewPathRes = NewPathRes + "\"
    mPathRes = NewPathRes
  End Property

  Property Get Language()
    Language = mLanguage
  End Property

  Sub SelectLocale(Default, Names, CookieExpired)
    Dim strLocale : strLocale = Empty
    Me.DefaultLanguage = Default
    If UseStaticTranslation Then 
      Set Locale=CCCreateLocale(GetText("CCS_FormatInfo",Empty))
      ChangeFormatInfo
      Exit Sub
    End If 
    LoadFormatInfo
    If Not IsEmpty(Names(CCS_I18N_RequestParameterName)) Then
      strLocale = FindLocale(Request.QueryString(Names(CCS_I18N_RequestParameterName)))
      If IsEmpty(strLocale) Then _
        strLocale = FindLocale(CCGetFromPost(Names(CCS_I18N_RequestParameterName), ""))
    End If
    If IsEmpty(strLocale) And Not IsEmpty(Names(CCS_I18N_SessionName)) Then _
      strLocale = FindLocale(Session(Names(CCS_I18N_SessionName)))
    If IsEmpty(strLocale) And Not IsEmpty(Names(CCS_I18N_LanguageSessionName)) Then _
      strLocale = FindLocale(Session(Names(CCS_I18N_LanguageSessionName)))
    If IsEmpty(strLocale) And Not IsEmpty(Names(CCS_I18N_CookieName)) Then _
      strLocale = FindLocale(Request.Cookies(Names(CCS_I18N_CookieName)))
    If IsEmpty(strLocale) And Not IsEmpty(Names(CCS_I18N_HttpHeaderName)) Then _
      strLocale = ParseHeader(Names(CCS_I18N_HttpHeaderName))
    If IsEmpty(strLocale) Then _
      strLocale = Default
    mLanguage = strLocale
    If Not Locales.Exists(mLanguage) Then _
      mLanguage = Default
    If Len(mLanguage) > 2 Then
       MainLanguage = Mid(Language, 1, 2)
    Else 
       MainLanguage = Empty
       Dim FullLocale : FullLocale = mLanguage & "-" & Locales(mLanguage)(2)
       If Locales.Exists(FullLocale) Then 
          MainLanguage = mLanguage
          mLanguage =  FullLocale
       End If
    End If

    If Not IsEmpty(Names(CCS_I18N_CookieName)) Then 
       If Request.Cookies(Names(CCS_I18N_CookieName)) <> mLanguage Then 
         Response.Cookies(Names(CCS_I18N_CookieName)) = mLanguage
         If Not IsEmpty(CookieExpired) Then _
           Response.Cookies(Names(CCS_I18N_CookieName)).Expires = DateAdd("d", CookieExpired, Now())
       End If
    End If
    If Not IsEmpty(Names(CCS_I18N_SessionName)) Then _
       Session(Names(CCS_I18N_SessionName)) = mLanguage
    If Not IsEmpty(Names(CCS_I18N_LanguageSessionName)) Then _
       Session(Names(CCS_I18N_LanguageSessionName)) = IIF(Len(MainLanguage) > 0, MainLanguage, mLanguage)
    Dim dum : dum = GetText("CCS_FormatInfo", Empty)
    Session.LCID = Locale.LCID 
  End Sub

  Function FindLocale(Language)
    Dim Lang, Country
    If Len(Language) >= 2 Then _
      Lang = LCase(Mid(Language, 1, 2))
    If Len(Language) >= 5 Then _
      Country = UCase(Mid(Language, 4, 2))
    If IsEmpty(Lang) Then
      FindLocale = Empty
    ElseIf IsEmpty(Country) Then
      If Locales.Exists(Lang) Then _
     FindLocale = Lang _
      Else _
     FindLocale = DefaultLanguage
    Else
      If Locales.Exists(Lang & "-" & Country) Then 
        FindLocale = Lang & "-" & Country 
      ElseIf Locales.Exists(Lang) Then
        FindLocale = Lang 
      Else 
        FindLocale = DefaultLanguage
      End If
    End If
  End Function

  Function ParseHeader(Value)
    Dim Poz, Langs, I, Ind 
    Dim Result : Result = False 
    Dim Lang 
    Lang = Request.ServerVariables(Value)
    Poz = InStr(Lang, ";")
    If Poz > 0 Then _
        Lang = Left(Lang, Poz)
    Langs = Split(Lang, ",")
    For I =0 To UBound(Langs)
      Lang = FindLocale(Langs(I))
      If Not IsEmpty(Lang) Then 
        Exit For
      End If
    Next    
    ParseHeader = Lang
  End Function

  Function GetText(MsgID, Params)
    Dim Result, I, Ind
    Result = Empty
    Dim sMsgID : sMsgID = MsgID
    If UseStaticTranslation Then 
      Result = Translation(MsgID)
    Else 
      If Not IsTranslationsLoaded(mLanguage) Then LoadTranslations(mLanguage)
      Result = Translation(MsgID)
      If IsEmpty(Result) And IsFallBack Then 
        If Not IsEmpty(MainLanguage) Then
          Result=FindTranslation(MsgID,MainLanguage)
        End If
        If IsEmpty(Result) Then 
          Result=FindTranslation(MsgID,DefaultLanguage)
        End If
      End If
    End If

    If Not IsEmpty(Result) Then 
      Result = Replace(Replace(Result, "\n", vbCrLf), "\\", "\")
      If IsArray(Params) Then
        For I = 0 To UBound(Params)
          Result = Replace(Result, "{" & I & "}", Params(I))
        Next
      Else
        Result = Replace(Result, "{0}", Params)
      End If
    Else 
      Result = sMsgID
    End If
    GetText = Result
  End Function

  Private Sub DeleteTranslations
    Dim Key
    Application.Lock
    For Each Key in Application.Contents
      If InStr(AppPrefix,Key)=0 Then _
        Application.Contents.Remove(Key)
    Next
    Application.UnLock
  End Sub
  

  Private Function IsTranslationsLoaded(Lang)
    Dim Key, FSO, F, ULang, FilePath 
    Dim Result : Result = False
    ULang = UCase(Lang)
    If Prepared = ULang Then    
      Result = True
    Else
      CachePrefix = AppPrefix & ULang 
      For Each Key in Application.Contents
        If Key = CachePrefix Then
          Set FSO = Server.CreateObject("Scripting.FileSystemObject")
          FilePath =  GetFilePath(Lang)
          If FSO.FileExists(FilePath) Then
            If FSO.GetFile(FilePath).DateLastModified = Application(CachePrefix)(2) And CanUseCache Then 
              Keys = Application(CachePrefix)(0)
              Vals = Application(CachePrefix)(1)
              Prepared = ULang
              If Locale is Nothing Then _ 
                Set Locale=CCCreateLocale(Locales(Lang)) 
              Result = True
            End If
          End If
    Set FSO = Nothing
          Exit For
        End If
      Next
    End If
    IsTranslationsLoaded = Result
  End Function 
    
  Private Function Translation(MsgID)
    If UBound(Keys) > 0 Then 
      Translation = BinarySearch(MsgID, Keys, Vals, True)
    Else 
      Translation = Empty
    End If
  End Function 


  Private Function FindTranslation(MsgID, Lang)
   If Not IsTranslationsLoaded(Lang) Then LoadTranslations(Lang)
   FindTranslation = Translation(MsgID)
  End Function 


  Private Sub LoadTranslations(Lang)
    PrepareTranslation FileRead(Lang), Lang
    If mCache Then
      CachePrefix = AppPrefix & UCase(Lang)   
      Dim AplArr(3)
      Application.Lock()
      AplArr(0) = Keys
      AplArr(1) = Vals
      AplArr(2) = DateLastModified
      Application(CachePrefix) = AplArr
      Application.Unlock()
      Prepared = UCase(Lang)
      CanUseCache = True
      If Locale is Nothing Then _ 
         Set Locale=CCCreateLocale(Locales(Lang))
    End If
  End Sub

  Public Function SetKeyVals( K, V)
    UseStaticTranslation = True
    Keys = K
    Vals = V
  End Function

  Private Function GetFilePath(Lang)
    GetFilePath = PathRes & Lang & Ext
  End Function

  Private Function FileRead(Lang)
    Dim FSO, Strm
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    FilePath = GetFilePath(Lang)
    If Not FSO.FileExists(FilePath) Then
'          Err.Raise 4010, "Internationalization engine", "File " & FilePath & " not found." 
        FileRead = ""
        Exit Function
    End If
    Set Strm = Server.CreateObject("ADODB.Stream")
    Strm.Open
    Strm.Charset = "utf-8"
    Strm.LoadFromFile FilePath
    Dim FileContent : FileContent = Strm.ReadText(adReadAll)
    Strm.Close
    DateLastModified =FSO.GetFile(FilePath).DateLastModified
    Set FSO = Nothing
    Set Strm = Nothing
    FileRead = FileContent
  End Function 


  Private Sub PrepareTranslation(FileContent, Lang)
    Dim i, Ar, Arl, l
    Ar = Split(Replace(FileContent, vbCr, ""), vbLf)
    ReDim Preserve Ar(UBound(Ar) + 2)
    Ar(UBound(Ar) - 1) = "CCS_LocaleID=" & Lang
    Ar(UBound(Ar)) = "CCS_LanguageID=" & Left(Lang, 2)
    If InStr(FileContent,"CCS_FormatInfo")=0 Then 
       ReDim Preserve Ar(UBound(Ar) + 1)
       Ar(UBound(Ar)) = "CCS_FormatInfo=" & IIF(IsArray(Locales(Lang)), Join(Locales(Lang),"|"), Locales(Lang))
    End If
    If UBound(Ar) >= 1 Then _
      QSort Ar, LBound(Ar), UBound(Ar)
    ReDim Keys(UBound(Ar))
    ReDim Vals(UBound(Ar))
    l = -1
    For i = 0 To UBound(Ar)
      If Left(LTrim(Ar(i)), 1) <> "'"  And InStr(Ar(i), "=") > 1 Then 
        Arl = split(Ar(i), "=")
        l = l + 1
        Keys(l) = LCase(Arl(0))
        Vals(l) = Arl(1)
      End If
    Next
    ReDim Preserve Keys(l)
    ReDim Preserve Vals(l)
    ChangeFormatInfo
  End Sub 
  
  Private Function BinarySearch(Trg, Keys, Vals, SearchValue)
    Dim index,first,last,middle
    last = UBound(Keys)
    first = LBound(Keys)
    Trg = LCase(Trg)
    BinarySearch = Empty
    Do
      middle = (first + last) \ 2
      If Keys(middle) = Trg Then
        BinarySearch = IIF(SearchValue,Vals(middle), middle)
        Exit Do
      ElseIf StrComp(Keys(middle),trg,1)<0 Then
        first = middle + 1
      Else
        last = middle - 1
      End If
    Loop Until first > last
  End Function

  Property Let Cache(newValue)
    If mCache Then DeleteTranslations
    IsTranslationsLoad = False
    mCache = Cache
  End Property

  Property Get Cache()
    Cache = mCache
  End Property


  Sub QSort(strList, lLbound, lUbound)
    Dim strTemp 
    Dim strBuffer
    Dim lngCurLow 
    Dim lngCurHigh
    Dim lngCurMidpoint 
    
    lngCurLow = lLbound
    lngCurHigh = lUbound
    
    If lUbound <= lLbound Then Exit Sub ' Error!
    lngCurMidpoint = (lLbound + lUbound) \ 2 
    strTemp = strList(lngCurMidpoint) 
    Do While (lngCurLow <= lngCurHigh)
      Do While StrCompLeft(strList(lngCurLow), strTemp) < 0   'strList(lngCurLow) < strTemp
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUbound Then Exit Do
      Loop

      Do While StrCompLeft(strTemp, strList(lngCurHigh))<0    'strTemp < strList(lngCurHigh)
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLbound Then Exit Do
      Loop

      If (lngCurLow <= lngCurHigh) Then 
        strBuffer = strList(lngCurLow)
        strList(lngCurLow) = strList(lngCurHigh)
        strList(lngCurHigh) = strBuffer
        lngCurLow = lngCurLow + 1 
        lngCurHigh = lngCurHigh - 1
      End If
    Loop

    If lLbound < lngCurHigh Then ' Recurse if necessary 
        QSort strList, lLbound, lngCurHigh
    End If
    If lngCurLow < lUbound Then  ' Recurse if necessary 
        QSort strList, lngCurLow, lUbound
    End If
  End Sub

  Private Function StrCompLeft(l, r)
     Dim arl : arl = Split(l, "=", 2)
     Dim arp : arp = Split(r, "=", 2)
	 If UBound(arl)>0 And UBound(arp)>0 Then 
        StrCompLeft = StrComp(arl(0), arp(0), 1)
     Else 
        StrCompLeft = StrComp(l, r, 1)
	 End If
  End Function

 Private Sub LoadFormatInfo()
    Dim FSO, Strm, FormatInfoPath, FormatInfoPrefix, Ar, I
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    FormatInfoPrefix = AppPrefix & "FormatInfo"   
    FormatInfoPath = PathRes & "formatting.txt"
    If Not FSO.FileExists(FormatInfoPath) Then _
       Err.Raise 4010, "Internationalization engine", "File " & FormatInfoPath & " not found." 
    If IsArray(Application(FormatInfoPrefix)) Then 
      Ar = Application(FormatInfoPrefix)
      If FSO.GetFile(FormatInfoPath).DateLastModified = Ar(1) Then 
        For I=0 To UBound(Ar(0))
          Locales(Ar(0)(I)(0)) = Ar(0)(I)
        Next
        Exit Sub
      End If     
    End If 

    DeleteTranslations
    CanUseCache = False

    Dim ArInfo(),ArAppl(2)
    ReDim ArInfo(50)
    Set Strm = Server.CreateObject("ADODB.Stream")
    Strm.Open
    Strm.Charset = "utf-8"
    Strm.LoadFromFile FormatInfoPath
    Dim FileContent : FileContent = Strm.ReadText(adReadAll)
    Strm.Close
    ArAppl(1) = FSO.GetFile(FormatInfoPath).DateLastModified
    Set FSO = Nothing
    Set Strm = Nothing
    Ar = Split(Replace(FileContent, vbCr, ""), vbLf)
    Dim l : l = 0
    For I=0 To UBound(Ar)
      If Left(LTrim(Ar(i)), 1) <> "'"  And Trim(Ar(i)) <> "" And InStr(Ar(i), "|") > 1 Then 
        ArInfo(l) = Split(Ar(i), "|")
        Locales(ArInfo(L)(0)) = ArInfo(l)
        l = l + 1
        If l >= UBound(ArInfo) Then ReDim Preserve ArInfo(UBound(ArInfo)+20) 
      End If
    Next
    ReDim Preserve ArInfo(l-1)
    ArAppl(0) = ArInfo
    Application.Lock()
    Application(FormatInfoPrefix) = ArAppl
    Application.Unlock()
  End Sub 

  Public Sub ChangeFormatInfo
    Dim Info 
    Dim Ind : Ind = BinarySearch("CCS_FormatInfo",Keys,Vals,False)
      Dim FormatInfo : FormatInfo = Vals(Ind)
    If IsArray(FormatInfo) Then 
      Info = FormatInfo
    Else 
      Info=Split(FormatInfo, "|")
    End If 
    Info(11) = Replace(Info(11),"!","")
    Info(12) = Replace(Info(12),"!","")
    Info(13) = Replace(Info(13),"!","")
    Info(14) = Replace(Info(14),"!","")
    If UBound(Info) >=17 Then ReDim Preserve  Info(17)
    Vals(Ind) = Join(Info,"|")
  End Sub

 End Class


'End clsLocales Class

'clsCCSLocaleInfo @0-B3B73449
Class clsCCSLocaleInfo
  Public Name
  Public Language
  Public Country
  Public BooleanFormat
  Public DecimalDigits
  Public DecimalSeparator
  Public GroupSeparator
  Public ZeroFormat
  Public NullFormat
  Public MonthNames
  Public MonthShortNames
  Public WeekdayNames
  Public WeekdayShortNames
  Public WeekdayNarrowNames
  Public ShortDate
  Public LongDate
  Public ShortTime
  Public LongTime
  Public GeneralDate
  Public FirstWeekDay
  Public AMDesignator
  Public PMDesignator
  Public Charset
  Public CodePage
  Public OverrideNumberFormats
  Public OverrideDateFormats
  Public LCID
End Class

Function CCCreateLocale(s)
   Dim L,Info
   If Not IsArray(s) Then 
      Info=Split(s, "|")
   Else 
      Info = s
   End If 
   Set L = New clsCCSLocaleInfo
   L.Name = Info(0)
   L.Language = Info(1)
   L.Country = Info(2)
   L.BooleanFormat = Split(Info(3), ";")
   L.DecimalDigits = Info(4)
   L.DecimalSeparator = Info(5)
   L.GroupSeparator = Info(6)
   L.MonthNames = Split(Info(7), ";")
   L.MonthShortNames = Split(Info(8), ";")
   L.WeekdayNames = Split(Info(9), ";")
   L.WeekdayShortNames = Split(Info(10), ";") 
   L.ShortDate = Split(Info(11), "!")
   L.LongDate = Split(Info(12), "!")
   L.ShortTime = Split(Info(13), "!")
   L.LongTime = Split(Info(14), "!")
   L.GeneralDate = Split(Info(11) & "! !" & Info(14),"!") 
   L.FirstWeekDay = Info(15)
   L.AMDesignator = Info(16)
   L.PMDesignator = Info(17)
   If UBound(Info) >=19 Then  
     L.Charset = Info(18)
     L.CodePage = Info(19)
     L.OverrideNumberFormats = IIF(Info(20) = "1", True, False)
     L.OverrideDateFormats = IIF(Info(21) = "1", True, False)
     L.WeekdayNarrowNames = Split(Info(22), ";")
     L.LCID = Info(23)
   Else 
     L.OverrideNumberFormats = False
     L.OverrideDateFormats = False
   End If 
   Set CCCreateLocale = L
End Function 

'End clsCCSLocaleInfo

'clsStringBuffer @0-0A3F192B
Class clsStringBuffer
  Private incremetRate
  Private itemCount 
  Private items
  
  Private Sub Class_Initialize()
    incremetRate = 50
    itemCount = 0
    ReDim items(incremetRate)
  End Sub
  
  Public Sub Append(ByVal strValue)
    If itemCount > UBound(items) Then
      ReDim Preserve items(UBound(items) + incremetRate)
    End If
    
    items(itemCount) = strValue
    itemCount = itemCount + 1
  End Sub
  
  Public Function ToString() 
    ToString = Join(items, "")
  End Function

End Class
'End clsStringBuffer

'clsSection Class @0-9C292B6F
Class clsSection
  Private m_Visible
  Private m_Height

  Private Sub Class_Initialize()
    m_Visible = True
  End Sub

  Public Property Get Height
    Height = m_Height
  End Property

  Public Property Let Height(value)
  If Not Visible Then
    m_Height = 0
  Else
    m_Height = value
  End If
  End Property

  Public Property Get Visible
    Visible = m_Visible
  End Property

  Public Property Let Visible(value)
    m_Visible = value
  End Property


End Class
'End clsSection Class

'clsCCCalendarEvents Class @0-67A004CC
Class clsCCCalendarEvents
  Private mEvents
  Private Sub Class_Initialize()
    Set mEvents= Server.CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate()
    Set mEvents = Nothing
  End Sub

  Public Function Add(Event_date, Events)
    Dim arEvents
    If mEvents.Exists(Event_date) Then 
      Set arEvents =  mEvents(Event_date)
      arEvents.Add(Events)
    Else
      Set arEvents =  New clsCCDynArray
      arEvents.Add(Events)
      mEvents.Add  Event_date, arEvents
    End If
  End Function

  Public Function GetEvents(Event_date)
    If mEvents.Exists(Event_date) Then 
      Set GetEvents = mEvents(Event_date)
    Else
      Set GetEvents = Nothing
    End If
  End Function
End Class

'End clsCCCalendarEvents Class

'clsCCDynArray Class @0-77977B65
Class clsCCDynArray
  Private incremetRate
  Private itemCount 
  Public items
  
  Private Sub Class_Initialize()
    incremetRate = 50
    itemCount = 0
    ReDim items(incremetRate)
  End Sub
  
  Public Sub Add(Itm)
    If itemCount > UBound(items) Then
      ReDim Preserve items(UBound(items) + incremetRate)
    End If
    If itemCount = 0 Then 
      Set items(itemCount) = Itm
      itemCount = itemCount + 1
    Else 
      Insert Itm
    End If 
  End Sub

  Private Sub Insert(trg)
    Dim index,first,last,middle,i
    last = itemCount - 1
    first = 0
    Do
      middle = (first + last) \ 2
     If items(middle).Time = trg.Time Then
        Exit Do
     ElseIf items(middle).Time < trg.Time Then
        first = middle + 1
     Else
        last = middle - 1
     End If
    Loop Until first > last
    if first>last Then 
      middle=first
    Else 
       middle=last
    End If
    for i=itemCount-1 to middle step -1
        Set items(i+1)=items(i)
    next 
    itemCount=itemCount+1
    Set items(middle)=trg
  End Sub

  Public Property Get Count
     Count = itemCount
  End Property 
End Class

'End clsCCDynArray Class

'clsDynamicArray Class @0-F181684B
Class clsDynamicArray
  Private mData

  Private Sub Class_Initialize()
    Redim mData(0)
  End Sub

  Public Property Get Item(iPos)
    If iPos < LBound(mData) or iPos > UBound(mData) then
      Exit Property   
    End If
    Set Item = mData(iPos)
  End Property

  Public Function ToArray()
    ToArray = mData
  End Function

  Public Property Let Item(iPos, varValue)
    If iPos < LBound(mData) Then Exit Property
    If iPos > UBound(mData) Then
      Redim Preserve mData(iPos)
    End If
    Set mData(iPos) = varValue
  End Property

End Class

'End clsDynamicArray Class


%>
