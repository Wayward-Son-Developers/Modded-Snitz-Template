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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_header.asp" -->
<%

Dim status, info1, info2, fStatus
status = Application(strCookieURL & "down")
fStatus = request.form("status")
DMessage = request.Form("DownMessage")

if DMessage = "" then
	DMessage = Application(strCookieURL & "DownMessage")
end if

if status = "" then
	status = false
end if

if (not isEmpty(fStatus)) and (Session(strCookieURL & "Approval") = strAdminCode) then 
	if status then
		Application.lock
		Application(strCookieURL & "down") = false
		Application(strCookieURL & "DownMessage") = ""
		Application.unlock
		status = false
	else
		Application.lock
		Application(strCookieURL & "down") = true
		Application(strCookieURL & "DownMessage") = DMessage
		Application.unlock
		status = true
	end if
end if

if status then
	info1 = "down"
	info2 = "Start"
else
	info1 = "running"
	info2 = "Stop"
end if

if Session(strCookieURL & "Approval") = strAdminCode Then
	Response.Write	"<table class=""misc"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""secondnav"">" & vbNewLine & _
					getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Start/Stop&nbsp;the&nbsp;Forum</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine
	
	Response.Write	"<form action=""down.asp"" method=""post"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td>Start/Stop the Forum</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	
	If status Then
		Response.Write	"<tr class=""warning"">" & vbNewLine
	Else
		Response.Write	"<tr class=""ok"">" & vbNewLine
	End If
	
	Response.Write	"<td>The current status of the boards is <span class=""hlf"">" & info1 & "</span>.</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr><td>" & vbNewLine & _
					"<input type=""hidden"" name=""status"" value=""" & status & """>" & vbNewLine & _
					"<p>The message below will appear when the board is closed.</p>" & vbNewLine & _
					"<textarea cols=""80"" rows=""12"" name=""DownMessage"" wrap=""soft"">" & Application(strCookieURL & "DownMessage") & "</textarea></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""options"">" & vbNewLine & _
					"<td><button type=""submit"" name=""Submit"">" & info2 & " the board</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
else  
	if not Application(strCookieURL & "down") then 
		response.redirect("default.asp")
	end if
	
	Response.Write	"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td>&ldquo;" & strForumTitle & "&rdquo; is currently closed.</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & vbNewLine & _
					"<p>The Administrator has chosen to close this forum with the following reason:</p>" & vbNewLine & _
					"<div class=""callout"">" & Application(strCookieURL & "DownMessage") & "</div>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""options"">" & vbNewLine & _
					"<td><a href=""admin_login.asp?target=down.asp"">Administrator Login</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine
end if

WriteFooter
%>
