<%
' --------------------------------------------------------------------
' Filename     : i_Email.asp
' Purpose      : Include Email Functions
' Date Created : 6/28/2006
' Created By   : Ben Shimshak
' Updated On   : 9/21/06 - Ben Shimshak - Added capability of sending to cc addresses and bcc addresses and sending to multiple to addresses.
' Required     : i_Connection.asp, i_ProcessDynamic.asp
'
' Functions    :
'
' EmbedImagesInBody(passString, passURL) - Embeds images in the body of the email
' SendEmailTemplate(passSiteID, passEmailTemplateID, passEmailFromAddress, passEmailFromName, passEmailReplyToAddress, passEmailReplyToName, passEmailToAddress, passEmailCCAddress, passEmailBCCAddress) - Retrieves an EmailTemplateID and sends out the email based on the row in the EmailTemplate table with that ID. If any of the other parameters are not NULL (nor ""), then they are used instead of the ones retrieved from the template.
' --------------------------------------------------------------------
%>
<%

dim arrCIDFIleNames

Function EmbedImagesInBody(passString, passURL)
'Function Description: Embeds images in the body of the email

 '   Dim tempStartPos, tempEndPos, tempCounter, tempLinkString, tempNewEndPos
  '  Dim tempFileNameStartPos, tempFileName, tempNewBody
   ' Dim tempEAID As String, tempECEID As String, tempParameter As String
    'Dim tempLeftBlock, tempRightBlock, tempMiddleBlock
    
    'tempEAID = passEAID
    'tempECEID = passECEID
    'tempCode = UCase(GetRandomCode(4))
    
   ' tempEAID = Chr(Asc(Left(tempEAID, 1)) + 17) & Right(tempEAID, Len(tempEAID) - 1)
    
    'For x = 4 To Len(tempEAID) Step 3
    '    tempEAID = Left(tempEAID, x - 1) & Chr(Asc(Mid(tempEAID, x, 1)) + 17) & Right(tempEAID, Len(tempEAID) - x)
    'Next
    
    'tempECEID = Chr(Asc(Left(tempECEID, 1)) + 17) & Right(tempECEID, Len(tempECEID) - 1)
    
    'For x = 4 To Len(tempECEID) Step 3
    '    tempECEID = Left(tempECEID, x - 1) & Chr(Asc(Mid(tempECEID, x, 1)) + 17) & Right(tempECEID, Len(tempECEID) - x)
    'Next
    
    'tempParameter = Left(tempCode, 2) & tempEAID & "Z" & tempECEID & Right(tempCode, 2)
    
    'If Not IsEmpty(passAB) Then
    '    tempParameter = tempParameter & "-" & passAB
    'End If
    dim tempLinkString, tempString, tempStartPos, tempCounter, tempEndPos, tempNewEndPos, tempLeftBlock,tempRightBlock,tempMiddleBlock, tempFileName
    tempLinkString = "src="
    tempString = passString
    tempStartPos = InStr(LCase(tempString), LCase(tempLinkString))
    tempCounter = 10

    Do While tempStartPos <> 0

        'first quote after the img src=
        tempEndPos = InStr(tempStartPos, LCase(tempString), "c")
        tempNewEndPos = InStr(tempEndPos + 1, LCase(tempString), """>") - 2

        If tempEndPos <> tempStartPos + 2 Then
            Exit Do
        End If


        'everything in tempstring up until the first quote
        tempLeftBlock = Left(tempString, tempStartPos + 4)

		'everything in the middle
		tempMiddleBlock = Mid(tempString, tempStartPos + 5, tempNewEndPos - tempStartPos - 3)
        tempMiddleBlock = Replace(tempMiddleBlock, passURL, "cid:")
        
        If LCase(Right(tempMiddleBlock, Len(tempMiddleBlock) - 4)) <> "tracker.asp" Then
            tempFileName = tempFileName & ", " & Right(tempMiddleBlock, Len(tempMiddleBlock) - 4)
        End If
        'everything from second quote till end
        tempRightBlock = Right(tempString, Len(tempString) - tempNewEndPos - 1)

        'if a second quote is not found, exit the loop
        If tempNewEndPos = 0 Then
            Exit Do
        End If      
		
        tempString = tempLeftBlock & tempMiddleBlock & tempRightBlock       
        tempStartPos = InStr(tempStartPos + 3, LCase(tempString), LCase(tempLinkString))

    Loop

    If Len(tempFileName) > 0 Then
        tempFileName = Right(tempFileName, Len(tempFileName) - 2)
    End If    
    'tempString = Replace(tempString, "cid:tracker.asp", passURL & "tracker.asp")
    arrCIDFileNames = Split(tempFileName, ",")
    EmbedImagesInBody = tempString
    
End Function

Function SendEmailTemplate(passSiteID, passEmailTemplateID, passEmailFromAddress, passEmailFromName, passEmailReplyToAddress, passEmailReplyToName, passEmailToAddress, passEmailCCAddress, passEmailBCCAddress)
'Function Description: Retrieves an EmailTemplateID and sends out the email based on the row in the EmailTemplate table with that ID. If any of the other parameters are not null (nor ""), then they are used instead of the ones retrieved from the template.

	'On Error Resume next

	'get site settings
	dim Mail, tempEmailDefaultSQL, tempEmailSQL, rsEmailDefaultObj, rsEmailObj, intSiteID, tempHelo, tempDNS, tempSMTPAuthentication, tempSMTPServer, tempSMTPPort, tempSMTPUserName, tempSMTPPassword, tempBounce, tempFromAddress, tempFromName

	Set Mail = Server.CreateObject("Persits.MailSender") 
	tempEmailDefaultSQL = "SELECT * FROM EmailDefault INNER JOIN Site ON EmailDefaultSiteID = SiteID WHERE SiteID = " & passSiteID
    set rsEmailDefaultObj = objConn.Execute(tempEmailDefaultSQL)

	intSiteID = rsEmailDefaultObj("SiteID")
        
    tempHelo = rsEmailDefaultObj("EmailDefaultDNSHelo")
    tempDNS = rsEmailDefaultObj("EmailDefaultDNS")
  
    If Not tempDNS Then
        tempSMTPAuthentication = rsEmailDefaultObj("EmailDefaultSMTPAuthentication")
        tempSMTPServer = rsEmailDefaultObj("EmailDefaultSMTPServer")
        tempSMTPPort = rsEmailDefaultObj("EmailDefaultSMTPPort")
        'if SMTP authentication is enabled, then set the SMTP Username and password
        If tempSMTPAuthentication Then
            tempSMTPUserName = rsEmailDefaultObj("EmailDefaultSMTPUserName")
            tempSMTPPassword = rsEmailDefaultObj("EmailDefaultSMTPPassword")
        End If
    End If
    
	tempBounce = rsEmailDefaultObj("EmailDefaultBounceBackEmail")
    tempFromAddress = rsEmailDefaultObj("EmailDefaultFromAddress")
    tempFromName = rsEmailDefaultObj("EmailDefaultFromName")
    tempReplyToAddress = rsEmailDefaultObj("EmailDefaultReplyToAddress")
    tempReplyToName = rsEmailDefaultObj("EmailDefaultReplyToName")
   
    Set rsEmailDefaultObj = Nothing


	'get email settings
	dim strDynamic, tempBody, tempBodyTextOnly, tempSubject, tempEmbedImages, tempImagePath, tempImageVirtualPath, tempSiteURL, EmbedImages, tempReplyToAddress, tempReplytoName, tempToAddress, tempCCAddress, tempBCCAddress

	tempEmailSQL = "SELECT * FROM EmailTemplate WHERE EmailTemplateNickname = '" & Replace(passEmailTemplateID, "'", "''") & "' AND EmailTemplateSiteID = " & intSiteID
    set rsEmailObj = objConn.Execute(tempEmailSQL)

	strDynamic = rsEmailObj("EmailTemplateBodyDynamicContent")

    tempBody = GetProcessedDynamicContent(strDynamic, rsEmailObj("EmailTemplateBody"))
    tempBodyTextOnly = GetProcessedDynamicContent(strDynamic, rsEmailObj("EmailTemplateBodyTextOnly"))

	tempSubject = rsEmailObj("EmailTemplateSubject")
    tempEmbedImages = rsEmailObj("EmailTemplateEmbedImages")
    tempImagePath = rsEmailObj("EmailTemplateImagePhysicalPath")
    tempImageVirtualPath = rsEmailObj("EmailTemplateImagePath")
    tempSiteURL = rsEmailObj("EmailTemplateImageURL")

    If tempEmbedImages Then
        EmbedImages = True
    Else
        EmbedImages = False
    End If
    
	If Len(rsEmailObj("EmailTemplateBounceBackEmail")) > 0 Then
        tempBounce = rsEmailObj("EmailTemplateBounceBackEmail")
    End If


	If Len(passEmailFromAddress) > 0 Then
		tempFromAddress = passEmailFromAddress
	ElseIf Len(rsEmailObj("EmailTemplateFromAddress")) > 0 Then
        tempFromAddress = rsEmailObj("EmailTemplateFromAddress")
    End If


	If Len(passEmailFromName) > 0 Then
		tempFromName = passEmailFromName
    ElseIf Len(rsEmailObj("EmailTemplateFromName")) > 0 Then
        tempFromName = rsEmailObj("EmailTemplateFromName")
    End If


	If Len(passEmailReplyToAddress) > 0 Then
		tempReplyToAddress = passEmailReplyToAddress
	ElseIf Len(rsEmailObj("EmailTemplateReplyToAddress")) > 0 Then
        tempReplyToAddress = rsEmailObj("EmailTemplateReplyToAddress")
    End If
    

	If Len(passEmailReplyToName) > 0 Then
		tempReplyToName = passEmailReplyToName
	ElseIf Len(rsEmailObj("EmailTemplateReplyToName")) > 0 Then
        tempReplyToName = rsEmailObj("EmailTemplateReplyToName")
    End If
    
	If Len(passEmailToAddress) > 0 Then
		tempToAddress = passEmailToAddress
	Else
		tempToAddress = rsEmailObj("EmailTemplateToAddress")
	End If
    
	If Len(passEmailCCAddress) > 0 Then
		tempCCAddress = passEmailCCAddress
	Else
		tempCCAddress = rsEmailObj("EmailTemplateCCAddress")
	End If
    
	If Len(passEmailBCCAddress) > 0 Then
		tempBCCAddress = passEmailCCAddress
	Else
		tempBCCAddress = rsEmailObj("EmailTemplateBCCAddress")
	End If

	tempSubject = GetProcessedDynamicContent(strDynamic, tempSubject)
	tempFromAddress = GetProcessedDynamicContent(strDynamic, tempFromAddress)
	tempFromName = GetProcessedDynamicContent(strDynamic, tempFromName)
    tempReplyToAddress = GetProcessedDynamicContent(strDynamic, tempReplyToAddress)
    tempReplyToName = GetProcessedDynamicContent(strDynamic, tempReplyToName)
	tempToAddress = GetProcessedDynamicContent(strDynamic, tempToAddress)
	tempCCAddress = GetProcessedDynamicContent(strDynamic, tempCCAddress)
	tempBCCAddress = GetProcessedDynamicContent(strDynamic, tempBCCAddress)
    
    Set rsEmailObj = Nothing
    
 
	'embed images
    Dim tempImageUrl

	tempImageURL = tempSiteURL & tempImageVirtualPath

	If IsNull(tempImageURL) Then tempImageURL = ""
	
	If InStr(Mid(tempImageURL, 8, Len(tempImageURL)), "//") <> 0 Then
		tempImageURLtemp = Replace(Mid(tempImageURL, 8, Len(tempImageURL)), "//", "/")
		tempImageURL = Left(tempImageURL, 7) & tempImageURLtemp
	End If

	dim fso, arrCount, strImagePath, tempFileName
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	If EmbedImages then
		tempBody = EmbedImagesInBody(tempBody, tempImageURL)
		If IsArray(arrCIDFileNames) Then
			For arrCount = 0 To UBound(arrCIDFileNames)
				If arrCIDFileNames(arrCount) <> "" Then
					strImagePath = tempImagePath & "\"
					strImagePath = strImagePath & arrCIDFileNames(arrCount)
					If fso.FileExists(strImagePath) Then
						Mail.AddEmbeddedImage strImagePath, arrCIDFileNames(arrCount)
					End If
				End If
			Next
		End If
	End If
	
	Dim x, toAddressList, ccAddressList, bccAddressList

	With Mail
		.Helo = tempHelo
	    .AddReplyTo tempReplyToAddress, tempReplyToName
		.From = tempFromAddress
		.FromName = tempFromName

		toAddressList = Split(tempToAddress, ", ")
		For x = 0 To UBound(toAddressList)
			.AddAddress toAddressList(x)
		Next

		If tempCCAddress <> "" Then
			ccAddressList = Split(tempCCAddress, ", ")
			For x = 0 To UBound(ccAddressList)
				.AddCC ccAddressList(x)
			Next
		End If

		If tempBCCAddress <> "" Then
			bccAddressList = Split(tempBCCAddress, ", ")
			For x = 0 To UBound(bccAddressList)
				.AddBCC bccAddressList(x)
			Next
		End If

		.Subject = tempSubject
		.Body = tempBody
		.AltBody = tempBodyTextOnly
		.IsHTML = true
		
		If Not tempDNS Then
			.Host = tempSMTPServer
			.Username = tempSMTPUserName
			.Password = tempSMTPPassword
			.Port = tempSMTPPort
		End If
		.MailFrom = tempBounce
		.Queue = True
		.Send
		tempFileName = Split(.QueueFileName, ".")
		.Reset
	End With

'	response.write "INSERT INTO EmailLog (EmailLogDateTime, EmailLogFileName, EmailLogType, EmailLogAddressTo, EmailLogBody,EmailLogStatus,EmailLogStatusDescription) values('" & now() & "', '" & tempFileName(0) & "', '" & passType & "', '" & rsClient("ClientEmail") & "', '" & tempBody & "', 'QUEUED', 'QUEUED TO BE SENT')"

    'SiteTemplateConn.Execute ("INSERT INTO EmailLog (EmailLogDateTime, EmailLogFileName, EmailLogType, EmailLogAddressTo, EmailLogBody,EmailLogStatus,EmailLogStatusDescription) values('" & now() & "', '" & tempFileName(0) & "', '" & passType & "', '" & rsClient("ClientEmail") & "', '" & replace(tempBody , "'", "''") & "', 'QUEUED', 'QUEUED TO BE SENT')")

	'If Err.number <> 0 Then
		'Dim strError
		'strError = "Error: SendEasternUCEmail - " & Err.source & " - " & Err.description
		'Call EmailHandleError(strError)
	'End if	
End Function
%>