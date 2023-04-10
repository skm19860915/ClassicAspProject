<%

'Navigator Class @0-DF642779
Const tpSimple   = 00001
Const tpCentered = 00002
Const tpMoving   = 00003

Function CCCreateNavigator(Target, Name, FileName, NumberPages, NavigatorType)
  Dim Navigator
  Set Navigator = New clsNavigator
  Navigator.Init Target, Name, FileName, NumberPages, NavigatorType
  Set CCCreateNavigator = Navigator
End Function

Class clsNavigator
  Public ComponentName, CCSEvents
  Public Visible

  Public QueryString
  Public NavigatorBlock
  Public FirstOn, FirstOff, PrevOn, PrevOff, NextOn, NextOff, LastOn, LastOff
  Public Pages, PageOn, PageOff

  Dim TargetName
  Dim DataSource
  Dim PageNumber
  Dim FileName
  Dim NumberPages
  Dim NavigatorType
  Dim PagesCount

  Private CCSEventResult
        
  Private Sub Class_Initialize()
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set DataSource = Nothing
    Visible = True
  End Sub

  Sub Show(Template)
    Dim LastPage
    Dim BeginPage, EndPage, J

    QueryString = CCGetQueryString("QueryString", Array(TargetName & "Page", "ccsForm"))
    Set NavigatorBlock = Template.Block("Navigator " & ComponentName)

    With NavigatorBlock
      Set FirstOn = .Block("First_On")
      Set FirstOff = .Block("First_Off")
      Set PrevOn = .Block("Prev_On")
      Set PrevOff = .Block("Prev_Off")
      Set NextOn = .Block("Next_On")
      Set NextOff = .Block("Next_Off")
      Set LastOn = .Block("Last_On")
      Set LastOff = .Block("Last_Off")
      Set Pages = .Block("Pages")
    End With

    If Not FirstOn Is Nothing Then FirstOn.Clear
    If Not FirstOff Is Nothing Then FirstOff.Clear
    If Not PrevOn Is Nothing Then PrevOn.Clear
    If Not PrevOff Is Nothing Then PrevOff.Clear
    If Not NextOn Is Nothing Then NextOn.Clear
    If Not NextOff Is Nothing Then NextOff.Clear
    If Not LastOn Is Nothing Then LastOn.Clear
    If Not LastOff Is Nothing Then LastOff.Clear
    If Not Pages Is Nothing Then Pages.Clear

    If PageNumber < 1 Then PageNumber = 1
    
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
    If Not Visible Then Exit Sub
    
    If Not DataSource is Nothing Then
      LastPage = DataSource.PageCount
      If LastPage = 0 Then
        If Not DataSource.RecordSet.Eof Then LastPage = PageNumber + 1
        If LastPage = 0 Then LastPage = 1
      End If
    Else
      LastPage = PagesCount
    End If

    ' Parse First and Prev blocks
    If PageNumber <= 1 Then
      If NOT FirstOff IS Nothing Then FirstOff.Visible = True
      If NOT PrevOff IS Nothing Then PrevOff.Visible = True
    Else
      If NOT FirstOn IS Nothing Then 
        With FirstOn
          .Variable("First_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", "1")
          .Visible = True
        End With
      End If
      If NOT PrevOn IS Nothing Then
        With PrevOn
          .Variable("Prev_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", PageNumber - 1)
          .Visible = True
        End With
      End If
    End If

    If NavigatorType = tpSimple Then
      ' Set Page Number
      Set Pages = NavigatorBlock.Block("Pages")
      If Not Pages Is Nothing Then 
        Set PageOff = Pages.Block("Page_Off")
				If Not PageOff Is Nothing Then
					PageOff.Variable("Page_Number") = PageNumber
					PageOff.ParseTo True, Pages
				End If
			Else
        NavigatorBlock.Variable("Page_Number") = PageNumber
      End If
    ElseIf NavigatorType = tpCentered OR NavigatorType = tpMoving Then
      Set Pages = NavigatorBlock.Block("Pages")

      If NOT Pages IS Nothing Then
        Set PageOn = Pages.Block("Page_On")
        Set PageOff = Pages.Block("Page_Off")

        If NOT (PageOn IS Nothing OR PageOff IS Nothing) Then

          Select Case NavigatorType

            Case tpCentered
              BeginPage = PageNumber - (NumberPages - 1) \ 2
              If BeginPage < 1 Then BeginPage = 1
              EndPage = BeginPage + NumberPages - 1
              If EndPage > LastPage Then 
                BeginPage = BeginPage - EndPage + LastPage
                If BeginPage < 1 Then BeginPage = 1
                EndPage = LastPage
              End If
              For J = BeginPage To EndPage
                If CLng(J) = CLng(PageNumber) Then
                  With PageOff
                    .Variable("Page_Number") = J
                    .ParseTo True, Pages
                  End With
                Else
                  With PageOn
                    .Variable("Page_Number") = J
                    .Variable("Page_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", J)
                    .ParseTo True, Pages
                  End With
                End If
              Next

            Case tpMoving
              Dim GroupNumber, GroupFloat
              GroupFloat = PageNumber / NumberPages 
              GroupNumber = Int(GroupFloat)
              If GroupFloat > GroupNumber Then GroupNumber = GroupNumber + 1
              BeginPage = 1 + NumberPages * (GroupNumber - 1)
              EndPage = NumberPages * GroupNumber
              If BeginPage < 1 Then BeginPage = 1
              If EndPage > LastPage Then EndPage = LastPage
              If BeginPage > 1 Then
                With PageOn
                  .Variable("Page_Number") = "&lt;" & (BeginPage - 1)
                  .Variable("Page_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", BeginPage - 1)
                  .ParseTo True, Pages
                End With
              End If
              For J = BeginPage To EndPage
                If CLng(J) = CLng(PageNumber) Then
                  With PageOff
                    .Variable("Page_Number") = J
                    .ParseTo True, Pages
                  End With
                Else
                  With PageOn
                    .Variable("Page_Number") = J
                    .Variable("Page_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", J)
                    .ParseTo True, Pages
                  End With
                End If
              Next
              If EndPage < LastPage Then
                With PageOn
                  .Variable("Page_Number") = (EndPage + 1) & "&gt;"
                  .Variable("Page_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", EndPage + 1)
                  .ParseTo True, Pages
                End With
              End If
          End Select
        End If
			Else
        NavigatorBlock.Variable("Page_Number") = PageNumber
      End If
    End If

    ' Set Total Pages
    NavigatorBlock.Variable("Total_Pages") = LastPage

    ' Parse Last and Next blocks
    If CLng(PageNumber) >= CLng(LastPage) Then
      If NOT NextOff IS Nothing Then NextOff.Visible = True
      If NOT LastOff IS Nothing Then LastOff.Visible = True
    Else
      If NOT NextOn IS Nothing Then 
        With NextOn
          .Variable("Next_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", PageNumber + 1)
          .Visible = True
        End With
      End If
      If NOT LastOn IS Nothing Then 
        With LastOn
          .Variable("Last_URL") = FileName & "?" & CCAddParam(QueryString, TargetName & "Page", LastPage)
          .Visible = True
        End With
      End If
    End If

    NavigatorBlock.Visible = True
  End Sub

  Sub SetDataSource ( objDataSource )
    Set DataSource = objDataSource
  End Sub

  Sub Init ( Target, Name, NewFileName, NewNumberPages, NewNavigatorType )
    TargetName = Target
    ComponentName = Name
    FileName      = NewFileName
    NumberPages   = NewNumberPages
    NavigatorType = NewNavigatorType
    PageNumber = CCGetParam(TargetName & "Page", 1)
    If Not IsNumeric(PageNumber) AND Len(PageNumber) > 0 Then
      PageNumber = 1
    ElseIf Len(PageNumber) > 0 Then
      If PageNumber > 0 Then PageNumber = CInt(PageNumber) Else PageNumber = 1
    Else
      PageNumber = 1
    End If
  End Sub

  Private Sub Class_Terminate()
    Set DataSource = Nothing
  End Sub

End Class
'End Navigator Class


%>
