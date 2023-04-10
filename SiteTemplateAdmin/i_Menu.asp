<%
' --------------------------------------------------------------------
' Filename     : i_Menu.asp
' Purpose      : Creates the string containing the HTML for the menu
' Date Created : 6/28/2006
' Created By   : Ben Shimshak
' Updated On   : 
' Required     : None
'
' Functions    : None
' --------------------------------------------------------------------
%>
<%

'Added By Clive - Display DHTML menu
Dim DHTMLMenu
Dim menuPath

dim tempPageName

Dim tempPath,arrTempPath
tempPath = LCase(Request.ServerVariables("SCRIPT_NAME"))
arrTempPath = split(tempPath, "/")

tempPageName = Split(Request.ServerVariables("SCRIPT_NAME"), "/")

if Request.Querystring("menu") <> "0" then
		If Session("GroupID") = 50 then
			menuPath = "menu-KeylexAdmin"
		end if
		If Session("GroupID") = 40 then
			menuPath = "menu-ClientAdmin"
		end if
		If Session("GroupID") = 30 then
			menuPath = "menu-KeylexUser"
		end if
		If Session("GroupID") = 20 then
			menuPath = "menu-ClientUser"
		end if

		DHTMLMenu = "<iframe src=""SiteSelect.asp"" allowtransparency=""yes"" style=""border: none; position: absolute; top: 0px; right: 0px; z-index: 1000; height: 30px;""></iframe>" & vbCrLf
		DHTMLMenu = DHTMLMenu & "<!-- DHTML Menu Builder Loader Code START -->"
		DHTMLMenu = DHTMLMenu & "<div id=DMBRI style=""position:absolute;"">"
		DHTMLMenu = DHTMLMenu & "<img src=""" & menuPath & "/images/dmb_i.gif"" name=DMBImgFiles width=""1"" height=""1"" border=""0"" alt="""">"
		DHTMLMenu = DHTMLMenu & "<img src=""" & menuPath & "/dmb_m.gif"" name=DMBJSCode width=""1"" height=""1"" border=""0"" alt="""">"
		DHTMLMenu = DHTMLMenu & "</div>"
		DHTMLMenu = DHTMLMenu & "<script language=""JavaScript"" type=""text/javascript"">"
		DHTMLMenu = DHTMLMenu & "var rimPath=null;var rjsPath=null;var rPath2Root=null;function InitRelCode(){var iImg;var jImg;var tObj;if(!document.layers){iImg=document.images['DMBImgFiles'];jImg=document.images['DMBJSCode'];tObj=jImg;}else{tObj=document.layers['DMBRI'];if(tObj){iImg=tObj.document.images['DMBImgFiles'];jImg=tObj.document.images['DMBJSCode'];}}if(!tObj){window.setTimeout(""InitRelCode()"",700);return false;}rimPath=_gp(iImg.src);rjsPath=_gp(jImg.src);rPath2Root=rjsPath+""../"";return true;}function _purl(u){return xrep(xrep(u,""%%REP%%"",rPath2Root),""\\"",""/"");}function _fip(img){if(img.src.indexOf(""%%REL%%"")!=-1) img.src=rimPath+img.src.split(""%%REL%%"")[1];return img.src;}function _gp(p){return p.substr(0,p.lastIndexOf(""/"")+1);}function xrep(s,f,n){if(s) s=s.split(f).join(n);return s;}InitRelCode();"
		DHTMLMenu = DHTMLMenu & "</script>"
		DHTMLMenu = DHTMLMenu & "<script language=""JavaScript"" type=""text/javascript"">"
		DHTMLMenu = DHTMLMenu & "function Loadmenu() {if(!rjsPath){window.setTimeout(""Loadmenu()"", 10);return false;}var navVer = navigator.appVersion;"
		DHTMLMenu = DHTMLMenu & "if(navVer.substr(0,3) >= 4)"
		DHTMLMenu = DHTMLMenu & "if((navigator.appName==""Netscape"") && (parseInt(navigator.appVersion)==4)) {"
		DHTMLMenu = DHTMLMenu & "document.write('<' + 'script language=""JavaScript"" type=""text/javascript"" src=""' + rjsPath + 'nsmenu.js""><\/script\>');"
		DHTMLMenu = DHTMLMenu & "} else {"
		DHTMLMenu = DHTMLMenu & "document.write('<' + 'script language=""JavaScript"" type=""text/javascript"" src=""' + rjsPath + 'iemenu.js""><\/script\>');"
		DHTMLMenu = DHTMLMenu & "}return true;}Loadmenu();</script>"
		DHTMLMenu = DHTMLMenu & "<!-- DHTML Menu Builder Loader Code END --><br><br>"


'End Added By Clive - Display DHTML menu
end if

%>
