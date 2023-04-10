<%

'Sorter Class @0-83485B65

Function CCCreateSorter(SorterName, Parent, FileName)
  Dim Sorter
  Set Sorter = New clsSorter
  With Sorter
    .ComponentName = SorterName
    .FileName = FileName
    Set .Parent = Parent
  End With
  Set CCCreateSorter = Sorter
End Function

Class clsSorter

  Public ComponentName, CCSEvents

  Dim OrderDirection
  Dim TargetName
  Dim FileName
  Dim Visible

  Private mParent
  Private CCSEventResult

  Private Sub Class_Initialize()
    Visible = True
    Set mParent = Nothing
    Set CCSEvents = CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate()
    Set mParent = Nothing
  End Sub

  Property Set Parent(newParent)
    Set mParent = newParent
  End Property

  Sub Show(Template)
    Dim IsOn, IsAsc
    Dim QueryString, SorterBlock
    Dim AscOnExist, AscOffExist, DescOnExist, DescOffExist
    Dim AscOn, AscOff, DescOn, DescOff

    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)

    If NOT Visible Then _
      Exit Sub

    TargetName = mParent.ComponentName
    IsOn = (mParent.ActiveSorter = ComponentName)
    IsAsc = (isEmpty(mParent.SortingDirection) OR mParent.SortingDirection = "ASC")

    Set SorterBlock = Template.Block("Sorter " & ComponentName)

    AscOnExist = SorterBlock.BlockExists("Asc_On", "block")
    If AscOnExist Then _
      Set AscOn = SorterBlock.Block("Asc_On")

    AscOffExist = SorterBlock.BlockExists("Asc_Off", "block")
    If AscOffExist Then _
      Set AscOff = SorterBlock.Block("Asc_Off")

    DescOnExist = SorterBlock.BlockExists("Desc_On", "block")
    If DescOnExist Then _
      Set DescOn = SorterBlock.Block("Desc_On")

    DescOffExist = SorterBlock.BlockExists("Desc_Off", "block")
    If DescOffExist Then _
      Set DescOff = SorterBlock.Block("Desc_Off")

    QueryString = CCGetQueryString("QueryString", Array(TargetName & "Page", "ccsForm"))
    QueryString = CCAddParam(QueryString, TargetName & "Order", ComponentName)

    If IsOn then
      If IsAsc then 
        OrderDirection = "DESC"
        If AscOnExist Then AscOn.Visible = True
        If AscOffExist Then AscOff.Visible = False
        If DescOnExist Then DescOn.Visible = False
        If DescOffExist Then
          DescOff.Variable("Desc_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Dir", OrderDirection)
          DescOff.Visible = True
        End If
      Else 
        OrderDirection = "ASC"
        If AscOnExist Then AscOn.Visible = False
        If AscOffExist Then 
          AscOff.Variable("Asc_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Dir", OrderDirection)
          AscOff.Visible = True
        End If
        If DescOnExist Then DescOn.Visible = True
        If DescOffExist Then DescOff.Visible = False
      End if
    Else
      OrderDirection = "ASC"
      If AscOnExist Then AscOn.Visible = False
      If AscOffExist Then 
        AscOff.Variable("Asc_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Dir", "ASC")
        AscOff.Visible = True
      End If
      If DescOnExist Then DescOn.Visible = False
      If DescOffExist Then 
        DescOff.Variable("Desc_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Dir", "DESC")
        DescOff.Visible = True
      End If
    End If

    QueryString = CCAddParam(QueryString, TargetName & "Dir", OrderDirection)
    SorterBlock.Variable("Sort_URL") = FileName & "?" & QueryString
    SorterBlock.Visible = True
  End Sub

End Class
'End Sorter Class


%>
