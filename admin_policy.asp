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
'## MOD: Custom Policy v1.2 for Snitz Forums v3.4
'## Author: Michael Reisinger (OneWayMule)
'## File: admin_policy.asp
'##
'## Get the latest version of this MOD at
'## http://www.onewaymule.org/onewayscripts/
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if

Response.Write  "<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Custom&nbsp;Policy</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	strPolicyMode = CLng(Request.Form("PolicyMode"))
	strPolicyContent = ChkString(Request.Form("Message"),"message")
	strsql = "UPDATE " & strTablePrefix & "CUSTOM_POLICY SET CP_MODE='" & strPolicyMode & "', CP_CONTENT='" & strPolicyContent & "' WHERE CP_ID=1"
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	Call OkMessage("Configuration Posted","admin_policy.asp","Back To Custom Policy Admin")
else
	strsql = "SELECT CP_MODE, CP_CONTENT FROM " & strTablePrefix & "CUSTOM_POLICY WHERE CP_ID=1"
	Set rsp = my_conn.execute(strsql)
	strPolicyMode = rsp("CP_MODE")
	strPolicyContent = rsp("CP_CONTENT")
	rsp.close
	set rsp = nothing
	Response.Write  "<script language=""JavaScript"" type=""text/javascript"" src=""inc_code.js""></script>" & vbNewLine & _
					"<form action=""admin_policy.asp"" method=""post"" name=""PostTopic"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<table class=""admin"" width=""100%"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Custom Policy Configuration</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Display Mode:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"Default:<input type=""radio"" class=""radio"" name=""PolicyMode"" value=""0""" & chkRadio(strPolicyMode,0,true) & ">&nbsp;" & vbNewLine & _
					"Custom:<input type=""radio"" class=""radio"" name=""PolicyMode"" value=""1""" & chkRadio(strPolicyMode,1,true) & ">" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
					%><!--#INCLUDE FILE="inc_post_buttons.asp"--><%
	Response.Write  "<tr>" & vbNewLine & _
					"<td class=""formlabel"">Content:&nbsp;<br /><br />" & vbNewLine
	
	if strAllowHTML = "1" then
		Response.Write  "* HTML is ON<br />" & vbNewLine
	else
		Response.Write  "* HTML is OFF<br />" & vbNewLine
	end if
	
	if strAllowForumCode = "1" then
		Response.Write  "* <a href=""JavaScript:openWindow6('pop_forum_code.asp')"">Forum Code</a> is ON<br />" & vbNewLine
	else
		Response.Write  "* Forum Code is OFF<br />" & vbNewLine
	end if
	
	Response.Write  "</td>" & vbNewLine & _
					"<td class=""formvalue""><textarea name=""Message"" cols=""50"" rows=""12"" wrap=""VIRTUAL"" style=""width:100%"" onselect=""storeCaret(this);"" onclick=""storeCaret(this);"" onkeyup=""storeCaret(this);"" onchange=""storeCaret(this);"">" & CleanCode(strPolicyContent) & "</textarea></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Placeholders:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<b>[adminemail]</b> - defines the position of the &ldquo;Contact the Administrator&rdquo; text and link to Contact page<br />" & vbNewLine & _
					"<b>[forumurl]</b> - defines the position of the URL of the forum" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2"">" & vbNewLine & _
					"<button type=""submit"" name=""Submit"">Submit Policy</button>&nbsp;" & _
					"<button type=""button"" name=""Preview"" onclick=""OpenPreview()"">Preview Policy</button>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
End If

WriteFooter
Response.End
%>
