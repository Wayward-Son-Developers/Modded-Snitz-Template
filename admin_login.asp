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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"Admin Login","") & "&nbsp;Admin&nbsp;Login</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

fName = strDBNTFUserName
fPassword = ChkString(Request.Form("Password"), "SQLString")

RequestMethod = Request.ServerVariables("Request_method")
strTarget = trim(chkString(request("target"),"SQLString"))

if RequestMethod = "POST" or strAuthType = "nt" Then
	strEncodedPassword = sha256("" & fPassword)

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & trim(fName) & "'"
	if strAuthType = "db" then	strSql = strSql & " AND M_PASSWORD = '" & trim(strEncodedPassword) & "'"
	strSql = strSql & " AND M_LEVEL = 3 AND M_STATUS = 1"
	
	Set dbRs = my_Conn.Execute(strSql)
		
	If not(dbRS.EOF) then 
		Session(strCookieURL & "Approval") = strAdminCode
		
		if strTarget = "" then
			strTarget = "admin_home.asp"
		end if
		If strAuthType = "db" Then
			Call OkMessage("Login was successful!",strTarget,"Click here to Continue")
			WriteFooter
		Else
			Response.Redirect strTarget
		End If
		
		Response.End
	else
		Call FailMessage("<li>You are not allowed access.<br />If you think you have reached this message in error, please try again.</li>",False)
	end if
end if

Response.Write	"<form action=""admin_login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"<input type=""hidden"" value=""" & strTarget & """ name=""target"">" & vbNewLine & _
				"<table class=""admin"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td colspan=""2"">Admin Login</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">UserName:&nbsp;</td>" & vbNewLine & _
				"<td class=""formvalue""><input type=""text"" name=""Name"" style=""width:150px;""></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">Password:&nbsp;</td>" & vbNewLine & _
				"<td class=""formvalue""><input type=""Password"" name=""Password"" style=""width:150px;""></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""options"" colspan=""2""><button type=""submit"" id=""Submit1"" name=""Submit1"">Login</button></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"</form>" & vbNewLine
WriteFooter %>
