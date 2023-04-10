<!-- #INCLUDE File="i_Connection.asp" -->
<!-- #INCLUDE File="i_ProcessDynamic.asp" -->
<!-- #INCLUDE File="i_Email.asp" -->
<%
	SendEmailTemplate Request.Form("SiteID"), Request.Form("EmailTemplateID"), Request.Form("FromAddress"), Request.Form("FromName"), Request.Form("ReplyToAddress"), Request.Form("ReplyToName"), Request.Form("ToAddress"), Request.Form("CCAddress"), Request.Form("BCCAddress")
%>

The email has been sent.