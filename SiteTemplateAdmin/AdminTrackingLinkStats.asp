<!-- #INCLUDE VIRTUAL="/i_Connection.asp" -->
<!-- #INCLUDE Virtual="/i_Menu.asp" -->

<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1252">
<title>AdminTrackingLinkStats</title>
<meta content="CodeCharge Studio 3.0.3.1" name="GENERATOR">
<link rel="stylesheet" type="text/css" href="Styles/Blueprint/Style.css">
<script type="text/javascript" src="Functions.js"></script>
<script type="text/javascript" src="DatePicker.js"></script>
<script type="text/javascript">
var myDatePicker = new Object();
myDatePicker.style = "Styles/Blueprint/Style.css";
myDatePicker.format = "ShortDate";

function checkDates(thisForm) {
	var emptyField;
	switch ('')
	{
		case thisForm.searchStartDate.value.trim():
			emptyField = thisForm.searchStartDate.title;
			break;
		
		case thisForm.searchEndDate.value.trim():
			emptyField = thisForm.searchEndDate.title;
			break;
	}

	if(emptyField) {
		alert("You must enter a date in the following field:\n - " + emptyField);
		return false;
	}

	var startDate = new Date(thisForm.searchStartDate.value);
	var endDate = new Date(thisForm.searchEndDate.value);

	if(startDate > endDate) {
		alert("The start date can't be later than the end date.");
		return false;
	}

	return true;
}

function validateDate(dateText) {
	/*
		This function is used to help check the dateText to see if it is a valid date.
	*/
	
	var myDate = new Date(dateText);
	if( (dateText.trim() != '') && isNaN(myDate.valueOf()) ) {
		return false;
	}
	return true;
}

function validateFieldType(fieldObject, fieldValidationType) {
	var thisForm = fieldObject.form;
	
	var fieldValue;
	if(fieldObject.getAttribute("labelShowing") == "true") {
		fieldValue = fieldObject.getAttribute("currentValue");
	} else {
		fieldValue = fieldObject.value;
	}
	
	var passedValidation;
	switch(fieldValidationType) {
		case ('date'):
			passedValidation = validateDate(fieldValue);
			break;
		case ('email'):
			passedValidation = validateEmail(fieldValue);
			break;
		case ('number'):
			passedValidation = validateNumber(fieldValue);
			break;
		case ('integer'):
			passedValidation = validateInteger(fieldValue);
			break;
		default:
			alert('Invalid fieldCheckType passed to checkFieldType function.');
			return false;
	}
	
	if(!passedValidation) {
		fieldObject.value = '';
		
		alert('You must enter a valid ' + fieldValidationType + ' in the following field:\n - ' + fieldObject.title);
		
		return false;
	}
	
	return true;
}

/* String Functions */
String.prototype.trim = function() {
	//trims whitespace from the beginning and end of a string
	return this.replace(/^\s+/,'').replace(/\s+$/,'');
}

/* Error Functions */
Error.prototype.toString = function() {
	return this.message;
}
</script>
</head>
<body>
<%= DHTMLMenu %>

<%
	'Authenticate User
	Dim GroupList
	GroupList = "50,40"

	If Session("GroupID") = "" Or  Not (InStr("," & GroupList & ",", "," & Session("GroupID") & ",") > 0) Then
		Dim redirectPage, queryString
		redirectPage = Request.ServerVariables("SCRIPT_NAME")
		queryString = Request.ServerVariables("QUERY_STRING")
		If queryString <> "" Then queryString = "?" & queryString
		response.redirect "Login.asp?ret_link=" & Server.URLEncode(redirectPage & queryString) & "&type=notLogged"
	End If


	'declare variables needed
	Dim TrackingLinkGroupID, searchStartDate, searchEndDate, searchLinkType, TrackingLinkGroupSQL, rsTrackingLinkGroup, TrackingLinkCodeSQL, rsTrackingLinkCode, TotalVisits

	'set query string values
	TrackingLinkGroupID = CLng(request.queryString("TrackingLinkGroupID"))
	searchStartDate = request.queryString("searchStartDate")
	searchEndDate = request.queryString("searchEndDate")
	searchLinkType = request.queryString("searchLinkType")

	'set default search dates
	If searchStartDate = "" And searchEndDate = "" Then
		searchEndDate = FormatDateTime(Now(), 2)
		searchStartDate = FormatDateTime(DateAdd("m", -1, searchEndDate), 2)
	End If

	'get tracking link groups
	TrackingLinkGroupSQL = "SELECT TrackingLinkGroupID, TrackingLinkGroupName, "&_
			"("&_
			"SELECT COUNT(*) "&_
			"FROM TrackingLinkCode "&_
			"LEFT OUTER JOIN TrackingLinkLog ON TrackingLinkLogTrackingLinkCode = TrackingLinkCode AND TrackingLinkLogSiteID = TrackingLinkCodeSiteID "&_
			"WHERE TrackingLinkLogDateTime >= '" & searchStartDate & "' AND TrackingLinkLogDateTime < '" & FormatDateTime(DateAdd("d", 1, searchEndDate), 2) & "' "&_
			"AND TrackingLinkCodeTrackingLinkGroupID = tlg.TrackingLinkGroupID "&_
			"AND TrackingLinkCodeSiteID = '" & Session("SiteID") & "' "&_
			"AND ('" & searchLinkType & "' = '' OR TrackingLinkCodeTrackingType = '" & searchLinkType & "')"&_
			") TrackingLinkGroupVisits "&_
			"FROM TrackingLinkGroup tlg "&_
			"WHERE TrackingLinkGroupSiteID = '" & Session("SiteID") & "' "&_
			"AND TrackingLinkGroupID IN (SELECT TrackingLinkCodeTrackingLinkGroupID FROM TrackingLinkCode WHERE ('" & searchLinkType & "' = '' OR TrackingLinkCodeTrackingType = '" & searchLinkType & "'))"

	'response.write TrackingLinkGroupSQL
	Set rsTrackingLinkGroup = objConn.Execute(TrackingLinkGroupSQL)

	'get open tracking link group's codes
	If TrackingLinkGroupID <> "" Then
		TrackingLinkCodeSQL = "SELECT TrackingLinkCode, TrackingLinkCodeDescription, "&_
			"("&_
			"SELECT COUNT(*) "&_
			"FROM TrackingLinkLog "&_
			"WHERE TrackingLinkLogDateTime >= '" & searchStartDate & "' AND TrackingLinkLogDateTime < '" & FormatDateTime(DateAdd("d", 1, searchEndDate), 2) & "' "&_
			"AND TrackingLinkLogTrackingLinkCode = tlc.TrackingLinkCode "&_
			"AND TrackingLinkLogSiteID = '" & Session("SiteID") & "'"&_
			") TrackingLinkCodeVisits "&_
			"FROM TrackingLinkCode tlc "&_
			"WHERE TrackingLinkCodeTrackingLinkGroupID = '" & TrackingLinkGroupID & "' "&_
			"AND TrackingLinkCodeSiteID = '" & Session("SiteID") & "' "&_
			"AND ('" & searchLinkType & "' = '' OR TrackingLinkCodeTrackingType = '" & searchLinkType & "')"
		Set rsTrackingLinkCode = objConn.Execute(TrackingLinkCodeSQL)
	End If
%>

<form name="TrackingLinkLogSearch" onsubmit="return checkDates(this);">
  <table cellspacing="0" cellpadding="0" border="0" align="center">
    <tr>
      <td valign="top">
        <table class="Header" cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td class="HeaderLeft"><img src="Styles/Blueprint/Images/Spacer.gif" border="0"></td> 
            <th>Search Tracking Link Log </th>
 
            <td class="HeaderRight"><img src="Styles/Blueprint/Images/Spacer.gif" border="0"></td> 
          </tr>
 
        </table>
 
        <table class="Record" cellspacing="0" cellpadding="0">
          <tr class="Controls">
            <th>Start Date</th>
 
            <td><input maxlength="100" size="8" value="<%= searchStartDate %>" name="searchStartDate" onblur="validateFieldType(this, 'date');" title="Start Date">
              <a href="javascript:showDatePicker('myDatePicker','TrackingLinkLogSearch','searchStartDate');"><img src="Styles/Blueprint/Images/DatePicker.gif" border="0"></a></td> 
          </tr>
          <tr class="Controls">
            <th>End Date</th>
 
            <td><input maxlength="100" size="8" value="<%= searchEndDate %>" name="searchEndDate" onblur="validateFieldType(this, 'date');" title="End Date">
              <a href="javascript:showDatePicker('myDatePicker','TrackingLinkLogSearch','searchEndDate');"><img src="Styles/Blueprint/Images/DatePicker.gif" border="0"></a></td> 
          </tr>
          <tr class="Controls">
            <th>Type</th>
 
            <td><select name="searchLinkType" title="Tracking Link Type">
              <option value=""<% If searchLinkType = "" Then %> selected<% End If %>>ALL</option>
              <option value="0"<% If searchLinkType = "0" Then %> selected<% End If %>>Incoming</option>
              <option value="1"<% If searchLinkType = "1" Then %> selected<% End If %>>Outgoing</option>
              <option value="2"<% If searchLinkType = "2" Then %> selected<% End If %>>Internal</option>
			  </select></td> 
          </tr>
 
          <tr class="Bottom">
            <td align="right" colspan="2">
              <input type="submit" value="Search"></td> 
          </tr>
 
        </table>
 </td> 
    </tr>
 
  </table>
</form>

<br>

<table cellspacing="0" cellpadding="0" border="0" align="center" width="90%">
  <tr>
    <td valign="top">
      <table class="Header" cellspacing="0" cellpadding="0" border="0">
        <tr>
          <td class="HeaderLeft"><img src="Styles/Blueprint/Images/Spacer.gif" border="0"></td> 
          <th>List of Tracking Link Group </th>
 
          <td class="HeaderRight"><img src="Styles/Blueprint/Images/Spacer.gif" border="0"></td> 
        </tr>
 
      </table>
 
      <table class="Grid" cellspacing="0" cellpadding="0">
        <tr class="Caption">
          <th>Name</th>
 
          <th>Visits</th>
 
        </tr>
 
		<%
			Dim LinkHref, GroupNameColumnText

			LinkHref = Request.ServerVariables("SCRIPT_NAME") & "?searchStartDate=" & searchStartDate & "&searchEndDate=" & searchEndDate & "&searchLinkType=" & searchLinkType

			TotalVisits = 0

			While Not rsTrackingLinkGroup.EOF
				If CLng(rsTrackingLinkGroup("TrackingLinkGroupID")) <> TrackingLinkGroupID Then
					GroupNameColumnText = "<a href=""" & LinkHref & "&TrackingLinkGroupID=" & rsTrackingLinkGroup("TrackingLinkGroupID") & """ style=""color: blue;""><img src=""images/plus.gif"" border=""0"" />" & rsTrackingLinkGroup("TrackingLinkGroupName") & "</a>"
				Else
					GroupNameColumnText = "<a href=""" & LinkHref & """ style=""color: blue;""><img src=""images/minus.gif"" border=""0"" />" & rsTrackingLinkGroup("TrackingLinkGroupName") & "</a>"
				End If

				TotalVisits = TotalVisits + rsTrackingLinkGroup("TrackingLinkGroupVisits")
		%>

		<tr class="Row">
		  <td><%= GroupNameColumnText  %>&nbsp;</td> 
		  <td><%= rsTrackingLinkGroup("TrackingLinkGroupVisits") %>&nbsp;</td>
		</tr>
		<tr class="Separator">
		  <td colspan="3"><img src="Styles/Blueprint/Images/Spacer.gif" border="0"></td> 
		</tr>

		<%
				If CLng(rsTrackingLinkGroup("TrackingLinkGroupID")) = TrackingLinkGroupID Then
					While Not rsTrackingLinkCode.EOF
		%>

		<tr class="AltRow">
		  <td><img src="images/hr_l.gif" />&nbsp;<%= rsTrackingLinkCode("TrackingLinkCodeDescription") %> [<%= rsTrackingLinkCode("TrackingLinkCode") %>]&nbsp;</td> 
		  <td><img src="images/hr_l.gif" />&nbsp;<%= rsTrackingLinkCode("TrackingLinkCodeVisits") %>&nbsp;</td>
		</tr>
		<tr class="Separator">
		  <td colspan="3"><img src="Styles/Blueprint/Images/Spacer.gif" border="0"></td> 
		</tr>

		
		<%
						rsTrackingLinkCode.movenext
					Wend
				End If
				rsTrackingLinkGroup.movenext
			Wend
		%>

        <tr class="Caption">
          <th>Total Visits</th>
 
          <td><%= TotalVisits %></td>
 
        </tr>
      </table>
 </td> 
  </tr>
</table>
<br>
<br>
</body>
</html>