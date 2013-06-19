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
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;E-mail&nbsp;Server&nbsp;Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Request.Form("strMailServer") = "" and Request.Form("strMailMode") <> "cdonts" and Request.Form("strEmail") = "1" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the Address of your Mail Server</li>"
	end if
	if ((lcase(left(Request.Form("strMailServer"), 7)) = "http://") or (lcase(left(Request.Form("strMailServer"), 8)) = "https://")) and Request.Form("strEmail") = "1" then
		Err_Msg = Err_Msg & "<li>Do not prefix the Mail Server Address with http://, https:// or file://</li>"
	end if
	if Request.Form("strSender") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the E-mail Address of the Forum Administrator</li>"
	else
		if EmailField(Request.Form("strSender")) = 0 and Request.Form("strSender") <> "" then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid E-mail Address for the Forum Administrator</li>"
		end if
	end if
	if Request.Form("strRestrictReg") = 1 and Request.Form("strEmailVal") = 0 then
		Err_Msg = Err_Msg & "<li>Email Validation must be enabled in order to enable the Restrict Registration Option</li>"
	end if
	if not IsNumeric(Request.Form("intMaxPostsToEMail")) then
		Err_Msg = Err_Msg & "<li>Number of posts to allow sending e-mail must be a number</li>"
	end if

	if Err_Msg = "" then
		'## Forum_SQL
		for each key in Request.Form 
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next
		Application(strCookieURL & "ConfigLoaded") = ""

		Call OkMessage("Configuration Posted!","admin_home.asp","Back To Admin Home")
	else
		Call FailMessage(Err_Msg,True)
	end if
else
	Dim theComponent(20)
	Dim theComponentName(20)
	Dim theComponentValue(20)

	'## the components
	theComponent(0) = "ABMailer.Mailman"
	theComponent(1) = "Persits.MailSender"
	theComponent(2) = "SMTPsvg.Mailer"
	theComponent(3) = "SMTPsvg.Mailer"
	theComponent(4) = "CDONTS.NewMail"
	theComponent(5) = "CDONTS.NewMail"
	theComponent(6) = "CDO.Message"
	theComponent(7) = "dkQmail.Qmail"
	theComponent(8) = "Dundas.Mailer"
	theComponent(9) = "Dundas.Mailer"
	theComponent(10) = "Innoveda.MailSender"
	theComponent(11) = "Geocel.Mailer"
	theComponent(12) = "iismail.iismail.1"
	theComponent(13) = "Jmail.smtpmail"
	theComponent(14) = "Jmail.Message"
	theComponent(15) = "MDUserCom.MDUser"
	theComponent(16) = "ASPMail.ASPMailCtrl.1"
	theComponent(17) = "ocxQmail.ocxQmailCtrl.1"
	theComponent(18) = "SoftArtisans.SMTPMail"
	theComponent(19) = "SmtpMail.SmtpMail.1"
	theComponent(20) = "VSEmail.SMTPSendMail"

	'## the name of the components
	theComponentName(0) = "ABMailer v2.2+"
	theComponentName(1) = "ASPEMail"
	theComponentName(2) = "ASPMail"
	theComponentName(3) = "ASPQMail"
	theComponentName(4) = "CDONTS (IIS 3/4/5)"
	theComponentName(5) = "Chili!Mail (Chili!Soft ASP)"
	theComponentName(6) = "CDOSYS (IIS 5/5.1/6)"
	theComponentName(7) = "dkQMail"
	theComponentName(8) = "Dundas Mail (QuickSend)"
	theComponentName(9) = "Dundas Mail (SendMail)"
	theComponentName(10) = "FreeMailSender"
	theComponentName(11) = "GeoCel"
	theComponentName(12) = "IISMail"
	theComponentName(13) = "JMail 3.x"
	theComponentName(14) = "JMail 4.x"
	theComponentName(15) = "MDaemon"
	theComponentName(16) = "OCXMail"
	theComponentName(17) = "OCXQMail"
	theComponentName(18) = "SA-Smtp Mail"
	theComponentName(19) = "SMTP"
	theComponentName(20) = "VSEmail"

	'## the value of the components
	theComponentValue(0) = "abmailer"
	theComponentValue(1) = "aspemail"
	theComponentValue(2) = "aspmail"
	theComponentValue(3) = "aspqmail"
	theComponentValue(4) = "cdonts"
	theComponentValue(5) = "chilicdonts"
	theComponentValue(6) = "cdosys"
	theComponentValue(7) = "dkqmail"
	theComponentValue(8) = "dundasmailq"
	theComponentValue(9) = "dundasmails"
	theComponentValue(10) = "freemailsender"
	theComponentValue(11) = "geocel"
	theComponentValue(12) = "iismail"
	theComponentValue(13) = "jmail"
	theComponentValue(14) = "jmail4"
	theComponentValue(15) = "mdaemon"
	theComponentValue(16) = "ocxmail"
	theComponentValue(17) = "ocxqmail"
	theComponentValue(18) = "sasmtpmail"
	theComponentValue(19) = "smtp"
	theComponentValue(20) = "vsemail"

	Response.Write	"<form action=""admin_config_email.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">E-mail Server Configuration</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Select E-mail Component:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><select name=""strMailMode"">" & vbNewLine
	dim i, j
	j = 0
	for i=0 to UBound(theComponent)
		if IsObjInstalled(theComponent(i)) then 
			Response.Write	"<option value=""" & theComponentValue(i) & """" & chkSelect(strMailMode,theComponentValue(i)) & ">" & theComponentName(i) & "</option>" & vbNewline
		else
			j = j + 1
		end if
	next
	if j > UBound(theComponent) then
		Response.Write	"<option value=""None"">No Compatible Component Found</option>" & vbNewline
	end if 

	Response.Write	"</select>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#email')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">E-mail Mode:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strEmail"" value=""1"""
	if j > UBound(theComponent) then Response.Write(" disabled") else if lcase(strEmail) <> "0" then Response.Write(" checked")
	Response.Write	">&nbsp;" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strEmail"" value=""0"""
	if j > UBound(theComponent) then Response.Write(" checked") else if lcase(strEmail) = "0" then Response.Write(" checked")
	Response.Write	">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#email')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">E-mail Server Address:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strMailServer"" size=""40"" value=""" & strMailServer & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#mailserver')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Administrator E-mail Address:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strSender"" size=""40"" value=""" & strSender & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#sender')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Require Unique E-mail:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strUniqueEmail"" value=""1""" & chkRadio(strUniqueEmail,1,true) & ">&nbsp;" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strUniqueEmail"" value=""0""" & chkRadio(strUniqueEmail,1,false) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#UniqueEmail')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">E-mail Validation:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strEmailVal"" value=""1""" & chkRadio(strEmailVal,1,true) & ">&nbsp;" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strEmailVal"" value=""0""" & chkRadio(strEmailVal,1,false) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#EmailVal')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Filter known spam domains:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strFilterEMailAddresses"" value=""1""" & chkRadio(strFilterEMailAddresses,1,true) & ">&nbsp;" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strFilterEMailAddresses"" value=""0""" & chkRadio(strFilterEMailAddresses,1,false) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#EmailFilter')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Restrict Registration:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strRestrictReg"" value=""1""" & chkRadio(strRestrictReg,1,true) & ">&nbsp;" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strRestrictReg"" value=""0""" & chkRadio(strRestrictReg,1,false) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#RestrictReg')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Require Logon for sending Mail:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strLogonForMail"" value=""1""" & chkRadio(strLogonForMail,1,true) & ">&nbsp;" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strLogonForMail"" value=""0""" & chkRadio(strLogonForMail,1,false) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#LogonForMail')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Number of posts to allow sending e-mail:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""intMaxPostsToEMail"" size=""40"" maxlength=""10"" value=""" & intMaxPostsToEMail & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#MaxPostsToEMail')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Error if they don't have enough posts:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strNoMaxPostsToEMail"" size=""40"" maxlength=""255"" value=""" & strNoMaxPostsToEMail & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#NoMaxPostsToEMail')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"" id=""submit1"" name=""submit1"">Submit New Config</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if 
WriteFooter
Response.End

function IsObjInstalled(strClassString)
	on error resume next
	'## initialize default values
	IsObjInstalled = false
	Err.Clear
	'## testing code
	dim xTestObj
	set xTestObj = Server.CreateObject(strClassString)
	if Err.Number = 0 then
		IsObjInstalled = true
	end if
	'## cleanup
	set xTestObj = nothing
	Err.Clear
	on error goto 0
end function
%>
