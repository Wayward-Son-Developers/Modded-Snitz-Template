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
on error resume next
strName = my_Conn.Properties(0).name
strValue = my_Conn.Properties(0).value
on error goto 0

if Err.Number <> 0 then
	blnDisplay = False
else
	blnDisplay = True
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Server&nbsp;Information</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine
				
Response.Write	"<table class=""admin"">" & vbNewLine & _
				"<tr class=""warning"">" & vbNewLine & _
				"<td colspan=""2""><b>NOTE:</b> The following table will show you values of interest in setting up these forums. Most useful will be the line that shows the APPL_PHYSICAL_PATH. This can be used to properly write your DSN'less Connection String.</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td>Variable Name</td>" & vbNewLine & _
				"<td>Value</td>" & vbNewLine & _
				"</tr>" & vbNewLine

For Each key In Request.ServerVariables
	Response.Write	"<tr>" & vbNewLine & _
					"<td class=""formlabel"">" & key & ":&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">"
	
	If Request.ServerVariables(key) = "" Then
		Response.Write "&nbsp;"
	Else
		Response.Write Request.Servervariables(key)
	End If
	
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
Next

'Write out the VBScript Version ' style=""font-family:courier""
Response.Write	"<tr>" & vbNewLine & _
				"<td class=""formlabel"">Scripting Engine:&nbsp;</td>" & vbNewLine & _
				"<td class=""formvalue"">" & _
				ScriptEngine & " v" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & " build " & ScriptEngineBuildVersion & _
				"</td>" & vbNewLine & _
				"</tr>" & vbNewLine

If blnDisplay = True Then
	'## Code below added to show general ADO/Database Information
	Response.Write	"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Database Connection Properties</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	
	For Each item In my_Conn.Properties
		Response.Write	"<tr>" & vbNewLine & _
						"<td class=""formlabel"">" & item.name & ":&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">"
		
		If item.value = "" Then
			Response.Write	"&nbsp;"
		Else
			Response.Write	item.value
		End If
		
		Response.Write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine
	Next
	
	'## Code above added to show general ADO/Database Information
end if

Response.Write	"</table>" & vbNewLine

WriteFooter
Response.End
%>
