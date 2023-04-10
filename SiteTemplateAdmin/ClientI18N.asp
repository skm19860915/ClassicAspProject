<%
'Include Common Files @0-CBC343D5
%>
<!-- #INCLUDE FILE="Common.asp"-->
<!-- #INCLUDE FILE="Cache.asp" -->
<!-- #INCLUDE FILE="Template.asp" -->
<%
'End Include Common Files

'Main @0-0503ED9C
    Dim FilePath, FileSystem, FileName, Encoding, FileContent, LangID, Strm, Matches, Match, Nm
    Set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
    FileName = LCase(Request.QueryString("file"))
    If NOT ( FileName = "functions.js" OR FileName = "datepicker.js") Then
      Response.write " "
      Response.End
    End If
    FilePath = Server.MapPath(FileName)
    FileContent = ""
    If NOT (FileSystem Is Nothing OR FilePath = "") Then
      If FileSystem.FileExists(FilePath) Then
        Set Strm = Server.CreateObject("ADODB.Stream")
        Strm.Open
        Strm.Charset = "utf-8"
        Strm.LoadFromFile FilePath
        FileContent = Strm.ReadText(adReadAll)
        Strm.Close
        Set Strm = Nothing
      End If
    End If
    Set FileSystem = Nothing
    Dim RegExpObject
    Set RegExpObject = New RegExp
    RegExpObject.Pattern = "{res:(\w+)}"
    RegExpObject.IgnoreCase = True
    RegExpObject.Global = True
    Set Matches = RegExpObject.Execute(FileContent)
    For Each Match in Matches
       Nm = Mid(Match.Value, 6, Len(Match.Value) - 6)
       FileContent = Replace(FileContent, Match.Value, CCSLocales.GetText(Nm, Empty))
'       FileContent = Replace(FileContent, Match.Value, CCSLocales.GetText(Match.SubMatches(0),Empty))
    Next
    Set Matches = Nothing
    Set RegExpObject = Nothing
    Session.CodePage = 65001
    Response.AddHeader "Content-type", "text/javascript; charset=utf-8"
    Response.write FileContent         
    If InputCodePage <> "" Then _
      Session.CodePage = InputCodePage

'End Main

%>
