<%
' --------------------------------------------------------------------
' Filename     : i_ProcessDynamic.asp
' Purpose      : Include Processing Dynamic Content Functions
' Date Created : 6/28/2006
' Created By   : Ben Shimshak
' Updated On   : 
' Required     : None
'
' Functions    :
'
' getProcessedDynamicContent(strDynamic, strOriginalContent) - Returns the dynamicly replaced string of strOriginalContent with things in stryDynamic
' --------------------------------------------------------------------
%>
<%
Function getProcessedDynamicContent(strDynamic, strOriginalContent)
'Function Description: Returns the dynamicly replaced string of strOriginalContent with things in stryDynamic
	if len(strDynamic) > 0 then 
		strDynamic = replace(strDynamic, "", """")
		execute(strDynamic)	'FIRSTNAME = "Clive"
		dim arrDyanmic
		arrDynamic = split(strParse, ",") 'strParse = "FIRSTNAME,LASTNAME,EMAIL"
		for each x in arrDynamic
			strOriginalContent = Replace(strOriginalContent, "[" & x & "]", eval(x))
		next
	end If

	getProcessedDynamicContent = strOriginalContent
End Function
%>