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
<%
if MemberID > 0 then Response.Redirect("default.asp")
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"All Forums","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"Forum Login","") & "&nbsp;Member&nbsp;Login</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

fName = strDBNTFUserName
fPassword = ChkString(Request.Form("Password"), "SQLString")

RequestMethod = Request.ServerVariables("Request_Method")
strTarget = trim(chkString(request("target"),"SQLString"))

if RequestMethod = "POST" Then
	strEncodedPassword = sha256("" & fPassword)
	select case chkUser(fName, strEncodedPassword,-1)
		case 1, 2, 3, 4
			Call DoCookies(Request.Form("SavePassword"))
			strLoginStatus = 1
		case else
			strLoginStatus = 0
	end select

	If strLoginStatus = 1 Then
		if strTarget = "" then
			Call OkMessage("Login was successful!","default.asp","Click here to Continue")
		else
			Call OkMessage("Login was successful!",strTarget,"Click here to Continue")
		end if
		WriteFooter
		Response.End
	ElseIf strLoginStatus = 0 Then
		Call FailMessage("Your username and/or password was incorrect.",False)
	End If
end if

Response.Write	"<script language=""JavaScript"" type=""text/javascript"">document.Form1.Name.focus();</script>" & vbNewLine & _
				"<form action=""login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"<input type=""hidden"" value=""" & strTarget & """ name=""target"">" & vbNewLine & _
				"<table class=""admin"" width=""60%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td colspan=""2"">Member Login</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr class=""section"">" & vbNewLine & _
				"<td width=""50%"">Login:</td>" & vbNewLine & _
				"<td width=""50%"">Login Questions:</td>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td>Username:<br />" & vbNewLine & _
				"<input type=""text"" name=""Name"" size=""20"" maxLength=""25"" tabindex=""1"" value="""" style=""width:150px;""><br />" & vbNewLine & _
				"Password:<br />" & vbNewLine & _
				"<input type=""password"" name=""Password"" size=""20"" tabindex=""2"" maxLength=""25"" value="""" style=""width:150px;""><br />" & vbNewLine & _
				"<input type=""checkbox"" name=""SavePassWord"" tabindex=""4"" value=""true"" checked>&nbsp;Save Password<br /><br />" & vbNewLine

if strGfxButtons = "1" then
	Response.Write	"<input src=""" & strImageUrl & "button_login.gif"" type=""image"" border=""0"" value=""Login"" id=""submit1"" name=""submit1"" tabindex=""3"">" & vbNewLine
else
	Response.Write	"<button type=""submit"" id=""submit1"" name=""submit1"" tabindex=""3"">Login</button>" & vbNewLine
end if 

Response.Write	"<td>" & vbNewLine & _
				"<p><acronym title=""Do I have to register?""><a href=""faq.asp#register""" & dWStatus("Do I have to register?") & ">Do I have to register?</a></acronym></p>" & vbNewLine
if strEmail = "1" then Response.Write("<p><acronym title=""Choose a new password if you have forgotten your current one.""><a href=""password.asp""" & dWStatus("Choose a new password if you have forgotten your current one.") & ">Forgot your Password?</a></acronym></p>" & vbNewLine)
Response.Write	"<p>Not a member? "
if strProhibitNewMembers = "1" then
	Response.Write	"<span class=""hlf"">The Administrator has turned off Registration for this forum. Only registered members are able to log in.</span>" & vbNewLine
else
	Response.Write	"<acronym title=""Click here to register.""><a href=""register.asp""" & dWStatus("Click here to register.") & ">Register Here!</a></acronymn>" & vbNewLine
end if

Response.Write	"</p></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"</form>" & vbNewLine
WriteFooter
%>
