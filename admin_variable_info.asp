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
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Forum&nbsp;Variables&nbsp;Information</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

Response.Write	"<table class=""admin"">" & vbNewLine & _
				"<tr class=""warning"">" & vbNewLine & _
				"<td colspan=""2""><b>NOTE:</b> The following table will show you values of the different variables used by the Forum.</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td>Variable&nbsp;Name</td>" & vbNewLine & _
				"<td>Value</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr class=""section"">" & vbNewLine & _
				"<td colspan=""2"">General&nbsp;information</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">strCookieUrl</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(StrCookieUrl, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">strUniqueID</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(StrUniqueID, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">strAuthType</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(strAuthType, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">strDBNTSQLName</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(strDBNTSQLName, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">strDBNTUserName</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(strDBNTUserName, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">strDBType</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(strDBType, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"">intCookieDuration</td>" & vbNewLine & _
				"<td class=""formvalue"">" & ChkString(intCookieDuration, "admindisplay") & "</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr class=""section"">" & vbNewLine & _
				"<td colspan=""2"">Cookies</td>" & vbNewLine & _
				"</tr>" & vbNewLine
				
for each key in Request.Cookies 
	if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
		if Request.Cookies(key).HasKeys then
			for each subkey in Request.Cookies(key)
				Response.Write	"<tr>" & vbNewLine & _
								"<td class=""formlabel"">" & chkString(key, "admindisplay") & " (" & chkString(subkey, "admindisplay") & ")</td>" & vbNewLine & _
								"<td class=""formvalue"">"
				if Request.Cookies(key)(subkey) = "" then
					Response.Write "&nbsp;"
				else
					Response.Write ChkString(CStr(Request.Cookies(key)(subkey)), "admindisplay")
				end if 
				Response.Write	"</td>" & vbNewline & _
								"</tr>" & vbNewline
			next
		else
			Response.Write	"<tr>" & vbNewLine & _
							"<td class=""formlabel"">" & chkString(key, "admindisplay") & "</td>" & vbNewLine & _
							"<td class=""formvalue"">"
			if Request.Cookies(key) = "" then
				Response.Write	"&nbsp;"
			else
				Response.Write	ChkString(CStr(Request.Cookies(key)), "admindisplay")
			end if 
			Response.Write	"</td>" & vbNewline & _
							"</tr>" & vbNewline
		end if
	end if
next

Response.Write	"<tr class=""section"">" & vbNewLine & _
				"<td colspan=""2"">Session&nbsp;variables</td>" & vbNewLine & _
				"</tr>" & vbNewLine

for each key in Session.Contents
	if not IsArray(Session.Contents(key)) then
		if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
			Response.Write	"<tr>" & vbNewLine & _
							"<td class=""formlabel"">" & ChkString(key, "admindisplay") & "</td>" & vbNewLine & _
							"<td class=""formvalue"">"
			if Session.Contents(key) = "" then
				Response.Write "&nbsp;"
			else
				Response.Write chkString(CStr(Session.Contents(key)), "admindisplay")
			end if 
			Response.Write	"</td>" & vbNewline & _
							"</tr>" & vbNewline
		end if
	end if
next 

Response.Write	"<tr class=""section"">" & vbNewLine & _
				"<td colspan=""2"">Application&nbsp;variables</td>" & vbNewLine & _
				"</tr>" & vbNewLine

for each key in Application.Contents
	if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
		Response.Write	"<tr>" & vbNewLine & _
						"<td class=""formlabel"">" & chkString(key, "admindisplay") & "</td>" & vbNewLine & _
						"<td class=""formvalue"">"
		if Application.Contents(key) = "" then
			Response.Write	"&nbsp;"
		else
			Response.Write	chkString(CStr(Application.Contents(key)), "admindisplay")
		end if 
		Response.Write	"</td>" & vbNewline & _
						"</tr>" & vbNewline
	end if
next 

Response.Write	"</table>" & vbNewline

WriteFooter
Response.End
%>
