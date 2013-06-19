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
<%

if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if

if strAuthType <> "nt" then
	Response.Redirect "admin_home.asp"
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Feature&nbsp;NT&nbsp;Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Err_Msg = "" then
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
	Response.Write	"<form action=""admin_config_NT_features.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Feature NT Configuration</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	if strAuthType = "nt" then
		Response.Write	"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Use NT Groups:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"On: <input type=""radio"" class=""radio"" name=""strNTGroups"" value=""1""" & chkRadio(strNTGroups,0,false) & ">&nbsp;" & vbNewLine & _
						"Off: <input type=""radio"" class=""radio"" name=""strNTGroups"" value=""0""" & chkRadio(strNTGroups,0,true) & ">" & vbNewLine & _
						"<a href=""JavaScript:openWindow3('pop_config_help.asp')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
		
		Response.Write	"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Use NT AutoLogon:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"On: <input type=""radio"" class=""radio"" name=""strAutoLogon"" value=""1""" & chkRadio(strAutoLogon,0,false) & ">&nbsp;" & vbNewLine & _
						"Off: <input type=""radio"" class=""radio"" name=""strAutoLogon"" value=""0""" & chkRadio(strAutoLogon,0,true) & ">" & vbNewLine & _
						"<a href=""JavaScript:openWindow3('pop_config_help.asp')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	end if
	Response.Write	"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"" id=""submit1"" name=""submit1"">Submit New Config</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if 
WriteFooter
Response.End
%>
