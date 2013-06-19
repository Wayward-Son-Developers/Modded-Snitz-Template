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
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<%
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Forgot your Password?</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if lcase(strEmail) <> "1" then
	Response.Redirect("default.asp")
end if

if Request.Form("mode") <> "DoIt" and Request.Form("mode") <> "UpdateIt" and trim(Request.QueryString("pwkey")) = "" then
	call ShowForm
elseif trim(Request.QueryString("pwkey")) <> "" and Request.Form("mode") <> "UpdateIt" then
	key = chkString(Request.QueryString("pwkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_PWKEY, MEMBER_ID, M_NAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_PWKEY = '" & key & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		Err_Msg = "<li>Your password key did not match the one that we have in our database.<br />Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br />If this problem persists, please <a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">contact the Administrator</a> of the forums.</li>"
		Call FailMessage(Err_Msg,False)
	elseif strComp(key,rsKey("M_PWKEY")) <> 0 then
		'Error message to user
		Err_Msg = "<li>Your password key did not match the one that we have in our database.<br />Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br />If this problem persists, please <a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">contact the Administrator</a> of the forums.</li>"
		Call FailMessage(Err_Msg,False)
	else
		PWMember_ID = rsKey("MEMBER_ID")
		call showForm2
	end if

	rsKey.close
	set rsKey = nothing
elseif trim(Request.Form("pwkey")) <> "" and Request.Form("mode") = "UpdateIt" then
	key = chkString(Request.Form("pwkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_PWKEY, MEMBER_ID, M_NAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))
	strSql = strSql & " AND M_PWKEY = '" & key & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		Err_Msg = "<li>Your password key did not match the one that we have in our database.<br />Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br />If this problem persists, please <a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">contact the Administrator</a> of the forums.</li>"
		Call FailMessage(Err_Msg,False)
	elseif strComp(key,rsKey("M_PWKEY")) <> 0 then
		'Error message to user
		Err_Msg = "<li>Your password key did not match the one that we have in our database.<br />Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br />If this problem persists, please <a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">contact the Administrator</a> of the forums.</li>"
		Call FailMessage(Err_Msg,False)
	else
		if trim(Request.Form("Password")) = "" then
			Err_Msg = Err_Msg & "<li>You must choose a Password</li>"
		end if
		if Len(Request.Form("Password")) > 25 then
			Err_Msg = Err_Msg & "<li>Your Password can not be greater than 25 characters</li>"
		end if
		if Request.Form("Password") <> Request.Form("Password2") then
			Err_Msg = Err_Msg & "<li>Your Passwords didn't match.</li>"
		end if

		if Err_Msg = "" then
			strEncodedPassword = sha256("" & Request.Form("Password"))

			'Update the user's password
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " SET M_PASSWORD = '" & chkString(strEncodedPassword,"SQLString") & "'"
			strSql = strSql & ", M_PWKEY = ''"
			strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))
			strSql = strSql & " AND M_PWKEY = '" & key & "'"

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		else
			Call FailMessage(Err_Msg,True)
			rsKey.close
			set rsKey = nothing
			WriteFooter
			Response.End
		end if
		
		if strAuthType = "db" then
			Call OkMessage("Your Password has been updated! You may now login with your UserName and new Password.","default.asp","Back To Forum")
		Else
			Call OkMessage("Your Password has been updated! You may now login.","default.asp","Back To Forum")
		End If
	end if

	rsKey.close
	set rsKey = nothing
else
	Err_Msg = ""

	if trim(Request.Form("Name")) = "" then
		Err_Msg = Err_Msg & "<li>You must enter your UserName</li>"
	end if

	if trim(Request.Form("Email")) = "" then
		Err_Msg = Err_Msg & "<li>You must enter your E-mail Address</li>"
	end if

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_NAME, M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & ChkString(Trim(Request.Form("Name")), "SQLString") &"'"
	strSql = strSql & " AND M_EMAIL = '" & ChkString(Trim(Request.Form("Email")), "SQLString") &"'"

	set rs = my_Conn.Execute (strSql)

	if rs.BOF and rs.EOF then
		Err_Msg = Err_Msg & "<li>Either the UserName or the E-mail Address you entered does not exist in the database.</li>"
	else
		PWMember_ID = rs("MEMBER_ID")
		PWMember_Name = rs("M_NAME")
		PWMember_Email = rs("M_EMAIL")
	end if
	
	rs.close
	set rs = nothing

	if Err_Msg = "" then
		pwkey = GetKey("none")

		'Update the user Member Level
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_PWKEY = '" & chkString(pwkey,"SQLString") & "'"
		strSql = strSql & " WHERE MEMBER_ID = " & PWMember_ID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		if lcase(strEmail) = "1" then
			'## E-mails Message to the Author of this Reply.  
			strRecipientsName = PWMember_Name
			strRecipients = PWMember_Email
			strFrom = strSender
			strFromName = strForumTitle
			strsubject = strForumTitle & " - Forgot Your Password? "
			strMessage = "Hello " & PWMember_Name & vbNewline & vbNewline
			strMessage = strMessage & "You received this message from " & strForumTitle & " because you have completed the First Step on the ""Forgot Your Password?"" page." & vbNewline & vbNewline
			strMessage = strMessage & "Please click on the link below to proceed to the next step." & vbNewline & vbNewLine
			strMessage = strMessage & strForumURL & "password.asp?pwkey=" & pwkey & vbNewline & vbNewline
			strMessage = strMessage & vbNewLine & "If you did not forget your password and received this e-mail in error, then you can just disregard/delete this e-mail, no further action is necessary." & vbNewLine & vbNewLine
			%><!--#INCLUDE FILE="inc_mail.asp" --><%
		end if
	else
		Call FailMessage(Err_Msg,True)
		WriteFooter
		Response.End
	end if
	Call OkMessage("Step One is Complete!</p><p>Please follow the instructions in the e-mail that has been sent to <b>" & ChkString(PWMember_Email,"email") & "</b> to complete the next step in this process.","default.asp","Back To Forum")
end if

WriteFooter
Response.End

sub ShowForm()
	Response.Write	"<form action=""password.asp"" method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input name=""mode"" type=""hidden"" value=""DoIt"">" & vbNewLine & _
					"<table class=""content"" width=""50%"">" & vbNewline & _
					"<tr class=""header"">" & vbNewline & _
					"<td colspan=""2"">Forgot your Password?</td>" & vbNewline & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewline & _
					"<td colspan=""2"">This is a 3 step process:" & vbNewLine & _
					"<ul>" & vbNewLine & _
					"<li><span class=""hlf""><b>First Step:</b> Enter your username and e-mail address in the form below to receive an e-mail containing a code to verify that you are who you say you are.</span></li>" & vbNewLine & _
					"<li><b>Second Step:</b> Check your e-mail and then click on the link that is provided to return to this page. Keep an eye on your spam filter as some programs will erroniously mark the e-mail as spam.</li>" & vbNewLine & _
					"<li><b>Third Step:</b> Choose your new password.</li>" & vbNewLine & _
					"</ul></td>" & vbNewline & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"" width=""50%"">UserName:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"" width=""50%""><input type=""text"" name=""Name"" size=""25"" maxLength=""25""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"" width=""50%"">E-mail Address:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"" width=""50%""><input type=""text"" name=""Email"" size=""25"" maxLength=""50""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"" id=""Submit1"" name=""Submit1"">Submit</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end sub

sub ShowForm2()
	Response.Write	"<form action=""password.asp"" method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input name=""mode"" type=""hidden"" value=""UpdateIt"">" & vbNewLine & _
					"<input name=""MEMBER_ID"" type=""hidden"" value=""" & PWMember_ID & """>" & vbNewLine & _
					"<input name=""pwkey"" type=""hidden"" value=""" & key & """>" & vbNewLine & _
					"<table class=""content"" width=""50%"">" & vbNewline & _
					"<tr class=""header"">" & vbNewline & _
					"<td colspan=""2"">Forgot your Password?</td>" & vbNewline & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewline & _
					"<td colspan=""2"">This is a 3 step process:" & vbNewLine & _
					"<ul>" & vbNewLine & _
					"<li><b>First Step:</b> Enter your username and e-mail address in the form below to receive an e-mail containing a code to verify that you are who you say you are.</li>" & vbNewLine & _
					"<li><b>Second Step:</b> Check your e-mail and then click on the link that is provided to return to this page. Keep an eye on your spam filter as some programs will erroniously mark the e-mail as spam.</li>" & vbNewLine & _
					"<li><span class=""hlf""><b>Third Step:</b> Choose your new password.</span></li>" & vbNewLine & _
					"</ul></td>" & vbNewline & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"" width=""50%"">Password:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"" width=""50%""><input name=""Password"" type=""Password"" size=""25"" maxLength=""25"" value=""""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"" width=""50%"">Password Again:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"" width=""50%""><input name=""Password2"" type=""Password"" maxLength=""25"" size=""25"" value=""""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"" id=""Submit1"" name=""Submit1"">Submit</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end sub
%>
