<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or (at your option) any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################

if trim(strFromName) = "" then
	strFromName = strForumTitle
end if

select case lcase(strMailMode) 
	case "abmailer"
		Set objNewMail = Server.CreateObject("ABMailer.Mailman")
		objNewMail.ServerAddr = strMailServer
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		objNewMail.SendTo = strRecipients
		objNewMail.MailSubject = strSubject
		objNewMail.MailMessage = strMessage
		on error resume next '## Ignore Errors
		objNewMail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "aspemail"
		Set objNewMail = Server.CreateObject("Persits.MailSender")
		objNewMail.FromName = strFromName
		objNewMail.From = strSender
		objNewMail.AddReplyTo strSender
		objNewMail.Host = strMailServer
		objNewMail.AddAddress strRecipients, strRecipientsName
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "aspmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		'objNewMail.AddReplyTo = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = strSubject
		objNewMail.BodyText = strMessage
		on error resume next '## Ignore Errors
		SendOk = objNewMail.SendMail
		If not(SendOk) <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & objNewMail.Response & "</li>"
		End if
	case "aspqmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		objNewMail.QMessage = 1
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = strSubject
		objNewMail.BodyText = strMessage
		on error resume next '## Ignore Errors
		objNewMail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "cdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		objNewMail.BodyFormat = 1
		objNewMail.MailFormat = 0
		on error resume next '## Ignore Errors
		objNewMail.Send strSender, strRecipients, strSubject, strMessage
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
		on error resume next '## Ignore Errors
	case "chilicdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		on error resume next '## Ignore Errors
		objNewMail.Host = strMailServer
		objNewMail.To = strRecipients
		objNewMail.From = strSender
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
		on error resume next '## Ignore Errors
	case "cdosys"
		Set iConf = Server.CreateObject ("CDO.Configuration")
		Set Flds = iConf.Fields 
		
		'Set and update fields properties
		Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
		Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer
		'Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
		Flds.Update
		
		Set objNewMail = Server.CreateObject("CDO.Message")
		Set objNewMail.Configuration = iConf
		
		'Format and send message
		Err.Clear 
		
		objNewMail.To = strRecipients
		objNewMail.From = strSender
		objNewMail.Subject = strSubject
		objNewMail.TextBody = strMessage
		On Error Resume Next
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "dkqmail"
		Set objNewMail = Server.CreateObject("dkQmail.Qmail")
		objNewMail.FromEmail = strSender
		objNewMail.ToEmail = strRecipients
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		objNewMail.CC = ""
		objNewMail.MessageType = "TEXT"
		on error resume next '## Ignore Errors
		objNewMail.SendMail()
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "dundasmailq"
		set objNewMail = Server.CreateObject("Dundas.Mailer")
		objNewMail.QuickSend strSender, strRecipients, strSubject, strMessage
		on error resume next '##Ignore Errors
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "dundasmails"
		set objNewMail = Server.CreateObject("Dundas.Mailer")
		objNewMail.TOs.Add strRecipients
		objNewMail.FromAddress = strSender
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '##Ignore Errors
		objNewMail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "freemailsender"
		set objNewMail = Server.CreateObject("Innoveda.MailSender")

		NoLoginMethod=0
		CramMD5Method=1
		AuthLoginMethod=2
		LoginPlainMethod=3

		objNewMail.Username = "username"
		objNewMail.Password = "password"
		objNewMail.LoginMethod = NoLoginMethod

		objNewMail.FromName = strFromName
		objNewMail.From = strSender
		'objNewMail.AddReplyTo strSender
		objNewMail.Host = strMailServer
		objNewMail.To = strRecipients
		'objNewMail.CC = strSender
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "geocel"
		set objNewMail = Server.CreateObject("Geocel.Mailer")
		objNewMail.AddServer strMailServer, 25
		objNewMail.AddRecipient strRecipients, strRecipientsName
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strFrom
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '##  Ignore Errors
		objNewMail.Send()
		if Err <> 0 then 
			Response.Write "Your request was not sent due to the following error: " & Err.Description 
		else
			Response.Write "Your mail has been sent..."
		end if
	case "iismail"
		Set objNewMail = Server.CreateObject("iismail.iismail.1")
		MailServer = strMailServer
		objNewMail.Server = strMailServer
		objNewMail.addRecipient(strRecipients)
		objNewMail.From = strSender
		objNewMail.Subject = strSubject
		objNewMail.body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "jmail"
		Set objNewMail = Server.CreateObject("Jmail.smtpmail")
		objNewMail.ServerAddress = strMailServer
		objNewMail.AddRecipient strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = strSubject
		objNewMail.body = strMessage
		objNewMail.priority = 3
		on error resume next '## Ignore Errors
		objNewMail.execute
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "jmail4"
		Set objNewMail = Server.CreateObject("Jmail.Message")
		'objNewMail.MailServerUserName = "myUserName"
		'objNewMail.MailServerPassword = "MyPassword"
		objNewMail.From = strSender
		objNewMail.FromName = strFromName
		objNewMail.AddRecipient strRecipients, strRecipientsName
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Send(strMailServer)
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "mdaemon"
		Set gMDUser = Server.CreateObject("MDUserCom.MDUser")
		mbDllLoaded = gMDUser.LoadUserDll
		if mbDllLoaded = False then
			response.write "Could not load MDUSER.DLL! Program will exit." & "<br />"
		else
			Set gMDMessageInfo = Server.CreateObject("MDUserCom.MDMessageInfo")
			gMDUser.InitMessageInfo gMDMessageInfo
			gMDMessageInfo.To = strRecipients
			gMDMessageInfo.From = strSender
			gMDMessageInfo.Subject = strSubject
			gMDMessageInfo.MessageBody = strMessage
			gMDMessageInfo.Priority = 0
			gMDUser.SpoolMessage gMDMessageInfo
			mbDllLoaded = gMDUser.FreeUserDll
		end if
		if Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		end if
	case "ocxmail"
		Set objNewMail = Server.CreateObject("ASPMail.ASPMailCtrl.1")
		recipient = strRecipients
		sender = strSender
		subject = strSubject
		message = strMessage
		mailserver = strMailServer
		on error resume next '## Ignore Errors
		result = objNewMail.SendMail(mailserver, recipient, sender, subject, message)
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "ocxqmail"
		Set objNewMail = Server.CreateObject("ocxQmail.ocxQmailCtrl.1")
		mailServer = strMailServer
		FromName = strFromName
		FromAddress = strSender
		priority = ""
		returnReceipt = ""
		toAddressList = strRecipients
		ccAddressList = ""
		bccAddressList = ""
		attachmentList = ""
		messageSubject = strSubject
		messageText = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Q mailServer,      _
			fromName,      _
		        fromAddress,      _
		        priority,      _
		        returnReceipt,      _
		        toAddressList,      _
		        ccAddressList,      _
		        bccAddressList,      _
		        attachmentList,      _
		        messageSubject,      _
		        messageText
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "sasmtpmail"
		Set objNewMail = Server.CreateObject("SoftArtisans.SMTPMail")
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		objNewMail.AddRecipient strRecipientsName, strRecipients
		'objNewMail.AddReplyTo strSender
		objNewMail.BodyText = strMessage
		objNewMail.organization = strForumTitle
		objNewMail.Subject = strSubject
		objNewMail.RemoteHost = strMailServer
		on error resume next
		SendOk = objNewMail.SendMail
		If not(SendOk) <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & objNewMail.Response & "</li>"
		End if
	case "smtp"
		Set objNewMail = Server.CreateObject("SmtpMail.SmtpMail.1")
		objNewMail.MailServer = strMailServer
		objNewMail.Recipients = strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = strSubject
		objNewMail.Message = strMessage
		on error resume next '## Ignore Errors
		objNewMail.SendMail2
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "vsemail"
		Set objNewMail = CreateObject("VSEmail.SMTPSendMail")
		objNewMail.Host = strMailServer
		objNewMail.From = strSender
		objNewMail.SendTo = strRecipients
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Connect
		objNewMail.Send
		objNewMail.Disconnect
		If Err <> 0 Then
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
end select

Set objNewMail = Nothing

on error goto 0
%>
