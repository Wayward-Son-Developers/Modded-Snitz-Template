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

'## Forum_SQL - Get membercount from DB 
strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS_PENDING"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSql, my_Conn

if not rs.EOF then
	User_Count = rs("U_COUNT")
else
	User_Count = 0
end if

rs.close
set rs = nothing

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Admin&nbsp;Section" & vbNewLine & _
				"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

select case strDBType
	case "access"
		if instr(lcase(strConnString), lcase(Server.MapPath("dbase/snitz_forums_2000.mdb"))) > 0 then
			Response.Write	"<div class=""oops"" style=""width:90%;"">" & vbNewLine & _
							"<p><b>WARNING:</b> The location of your access database may not be secure.</p>" & _
							"<p>You should consider moving the database from <b>" & _
							Server.MapPath("dbase/snitz_forums_2000.mdb") & _
							"</b> to a directory not directly accessible via a URL" & _
							" and/or renaming the database to another name.</p>" & _
							"<p><i>(After moving or renaming your database, remember to change the " & _
							"strConnString setting in config.asp.)</i></p>" & _
							"</div>" & vbNewLine
		end if
	case "sqlserver"
		if instr(lcase(strConnString), ";uid=sa;")> 0 then
			Response.Write	"<div class=""oops"" style=""width:90%;"">" & vbNewLine & _
							"<p><b>WARNING:</b> You are connecting to your MS SQL Server database with the <b>sa</b> user.</p>" & _
							"<p>After you have completed your installation, consider creating a new user with lower privileges" & _
							" and use that to connect to the database instead.</p>" & _
							"</div>" & vbNewLine
		end if
	case "mysql"
		if instr(lcase(strConnString), ";uid=root;")> 0 then
			Response.Write	"<div class=""oops"" style=""width:90%;"">" & vbNewLine & _
							"<p><b>WARNING:</b> You are connecting to your MySQL Server database with the <b>root</b> user.</p>" & _
							"<p>After you have completed your installation, consider creating a new user with lower privileges" & _
							" and use that to connect to the database instead.</p>" & _
							"</div>" & vbNewLine
		end if
end select

Response.Write	"<table class=""admin"" width=""100%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td colspan=""2"">Administrative Functions</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td width=""50%"">" & vbNewLine & _
				"<p>Forum Feature Configuration:" & vbNewLine & _
				"<ul>" & vbNewLine & _
				"<li><a href=""admin_config_system.asp"">Main Forum Configuration</a></li>" & vbNewLine & _
				"<li><a href=""admin_config_features.asp"">Feature Configuration</a></li>" & vbNewLine
if strAuthType = "nt" then
	Response.Write	"<li><a href=""admin_config_NT_features.asp"">Feature NT Configuration</a></li>" & vbNewLine
end if
Response.Write	"<li><a href=""admin_config_members.asp"">Member Details Configuration</a></li>" & vbNewLine & _
				"<li><a href=""admin_config_ranks.asp"">Ranking Configuration</a></li>" & vbNewLine & _
				"<li><a href=""admin_config_datetime.asp"">Server Date/Time Configuration</a></li>" & vbNewLine & _
				"<li><a href=""admin_config_email.asp"">Email Server Configuration</a></li>" & vbNewLine
If strFilterEMailAddresses = "1" Then Response.Write "<li><a href=""admin_spamserver.asp"">Blocked E-Mail Domains</a></li>" & vbNewLine
Response.Write	"<li><a href=""admin_config_colors.asp"">Color Code Configuration</a></li>" & vbNewLine & _
				"<li><a href=""javascript:openWindow3('admin_config_badwords.asp')"">Bad Word Filter Configuration</a></li>" & vbNewLine & _
				"<li><a href=""javascript:openWindow3('admin_config_namefilter.asp')"">UserName Filter Configuration</a></li>" & vbNewLine & _
				"<li><a href=""javascript:openWindow3('admin_config_order.asp')"">Category/Forum Order Configuration</a></li>" & vbNewLine & _
				"<li><a href=""admin_config_integration.asp"">Site Integration Configuration</a></li>" & vbNewLine & _
				"<li><a href=""admin_etc.asp"">Forum Cleanup Tools</a></li>" & vbNewLine & _
				"<li><a href=""admin_config_groupcats.asp"">Group Categories Configuration</a></li>" & vbNewLine & _
				"</ul></p>" & vbNewLine & _
				"</td>" & vbNewLine & _
				"<td width=""50%"">" & vbNewLine & _
				"<p>Other Configuration Options and Features:" & vbNewLine & _
				"<ul>" & vbNewLine
if strEmailVal = "1" then Response.Write("<li><a href=""admin_accounts_pending.asp"">Members Pending</a>&nbsp;<font size=""" & strFooterFontSize & """>(" & User_Count & ")</font></li>" & vbNewLine)
Response.Write  "<li><a href=""admin_usergroups.asp"">UserGroups Manager</a></li>" & vbNewLine
Response.Write	"<li><a href=""admin_members.asp"">Admin/Moderator List</a></li>" & vbNewLine & _
				"<li><a href=""admin_member_search.asp"">Member Search</a></li>" & vbNewLine & _
				"<li><a href=""admin_moderators.asp"">Moderator Setup</a></li>" & vbNewLine & _
				"<li><a href=""admin_emaillist.asp"">E-mail List</a></li>" & vbNewLine & _
				"<li><a href=""admin_policy.asp"">Custom Policy</a></li>" & vbNewLine & _
				"<li><a href=""admin_faq.asp"">F.A.Q. Administration</a></li>" & vbNewLine & _
				"<li><a href=""admin_info.asp"">Server Information</a></li>" & vbNewLine & _
				"<li><a href=""admin_variable_info.asp"">Forum Variables Information</a></li>" & vbNewLine & _
				"<li><a href=""admin_count.asp"">Update Forum Counts</a></li>" & vbNewLine
if strDBType = "access" and Instr(19,strConnString,"Jet",1) > 0 then Response.write("<li><a href=""admin_compactdb.asp"">Compact Database</a></li>" & vbNewLine)
if strArchiveState = "1" then Response.Write("<li><a href=""admin_forums.asp"">Forum Archive &amp; Cleanup</a></li>" & vbNewLine)
Response.Write	"<li><a href=""down.asp"">Start/Stop the Forum</a></li>" & vbNewLine & _
				"<li><a href=""admin_mod_dbsetup.asp"">MOD Setup</a>&nbsp;(<a href=""admin_mod_dbsetup2.asp"">Alternative MOD Setup</a>)</li>" & vbNewLine & _
				"<li><a href=""setup.asp"">Check Installation</a>&nbsp;<b>(Run after each upgrade !)</b></li>" & vbNewLine & _
				"</ul></p>" & vbNewLine & _
				"</td>" & vbNewLine & _
				"</tr>" & vbNewLine
Response.Write	"</table>" & vbNewLine
WriteFooter
Response.End
%>
