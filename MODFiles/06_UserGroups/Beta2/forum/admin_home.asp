<%
'#################################################################################
'## Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or any later version.
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
'## Support can be obtained from support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## reinhold@bigfoot.com
'##
'## or
'##
'## Snitz Communications
'## C/O: Michael Anderson
'## PO Box 200
'## Harpswell, ME 04079
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
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

select case strDBType
	case "access"
		if instr(strConnString, Server.MapPath("snitz_forums_2000.mdb"))> 0 then
			Response.Write	"    <br />" & vbNewLine & _
					"      <table border=""1"" width=""100%"" bgcolor=""red"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td align=""center""><font color=""white"" size=""2"">" & _
					"<b>WARNING:</b> The location of your access database may not be secure.<br /><br />" & _
					"You should consider moving the database from <b>" & Server.MapPath("snitz_forums_2000.mdb") & "</b> to a directory not directly accessable via a URL" & _
					" and/or renaming the database to another name." & _
					"<br /><br /><i>(After moving or renaming your database, remember to change the strConnString setting in config.asp.)</i>" & _
					"</font></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table><br />" & vbNewLine
		end if
	case "sqlserver"
		if instr(lcase(strConnString), ";uid=sa;")> 0 then
			Response.Write	"    <br />" & vbNewLine & _
					"      <table border=""1"" width=""100%"" bgcolor=""red"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td align=""center""><font color=""white"" size=""2"">" & _
					"<b>WARNING:</b> You are connecting to your MS SQL Server database with the <b>sa</b> user.<br /><br />" & _
					"After you have completed your installation, consider creating a new user with lower privileges" & _
					" and use that to connect to the database instead." & _
					"</font></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table><br />" & vbNewLine
		end if
	case "mysql"
		if instr(lcase(strConnString), ";uid=root;")> 0 then
			Response.Write	"    <br />" & vbNewLine & _
					"      <table border=""1"" width=""100%"" bgcolor=""red"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td align=""center""><font color=""white"" size=""2"">" & _
					"<b>WARNING:</b> You are connecting to your MySQL Server database with the <b>root</b> user.<br /><br />" & _
					"After you have completed your installation, consider creating a new user with lower privileges" & _
					" and use that to connect to the database instead." & _
					"</font></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table><br />" & vbNewLine
		end if
end select

Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Admin&nbsp;Section<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine & _
		"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Administrative Functions</b></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                <p><b>Forum Feature Configuration:</b>" & vbNewLine & _
		"                <UL>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_config_system.asp"">Main Forum Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_config_features.asp"">Feature Configuration</a></span></LI>" & vbNewLine
if strAuthType = "nt" then
	Response.Write	"                <LI><span class=""spnMessageText""><a href=""admin_config_NT_features.asp"">Feature NT Configuration</a></span></LI>" & vbNewLine
end if
Response.Write	"                <LI><span class=""spnMessageText""><a href=""admin_config_members.asp"">Member Details Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_config_ranks.asp"">Ranking Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_config_datetime.asp"">Server Date/Time Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_config_email.asp"">Email Server Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_config_colors.asp"">Font/Table Color Code Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""javascript:openWindow3('admin_config_badwords.asp')"">Bad Word Filter Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""javascript:openWindow3('admin_config_namefilter.asp')"">UserName Filter Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""javascript:openWindow3('admin_config_order.asp')"">Category/Forum Order Configuration</a></span></LI>" & vbNewLine & _
		"                </UL></p>" & vbNewLine & _
		"                </font></td>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                <p><b>Other Configuration Options and Features:</b>" & vbNewLine & _
		"                <UL>" & vbNewLine
if strEmailVal = "1" then Response.Write("                <LI><span class=""spnMessageText""><a href=""admin_accounts_pending.asp"">Members Pending</a></span>&nbsp;<font size=""" & strFooterFontSize & """>(" & User_Count & ")</font></LI>" & vbNewLine)
'#######    Begin UserGroup MOD     #######
Response.Write	"                <LI><span class=""spnMessageText""><a href=""admin_usergroups.asp"">UserGroups Manager</a></span></LI>" & vbNewLine
'#######     End UserGroup MOD      #######
Response.Write	"                <LI><span class=""spnMessageText""><a href=""admin_moderators.asp"">Moderator Setup</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_emaillist.asp"">E-mail List</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_info.asp"">Server Information</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_variable_info.asp"">Forum Variables Information</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_count.asp"">Update Forum Counts</a></span></LI>" & vbNewLine
if strArchiveState = "1" then Response.Write("                <LI><span class=""spnMessageText""><a href=""admin_forums.asp"">Archive Forum Topics</a></span></LI>" & vbNewLine)
Response.Write	"                <LI><span class=""spnMessageText""><a href=""admin_config_groupcats.asp"">Group Categories Configuration</a></span></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""admin_mod_dbsetup.asp"">MOD Setup</a></span><font size=""" & strFooterFontSize & """>&nbsp;(<span class=""spnMessageText""><a href=""admin_mod_dbsetup2.asp"">Alternative MOD Setup</a></span>)</font></LI>" & vbNewLine & _
		"                <LI><span class=""spnMessageText""><a href=""setup.asp"">Check Installation</a></span><font size=""" & strFooterFontSize & """><b> (Run after each upgrade !)</b></font></LI>" & vbNewLine & _
		"                </UL></p>" & vbNewLine & _
		"                </font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine
'Response.Write	"              <tr>" & vbNewLine & _
		'"                <td bgcolor=""" & strForumCellColor & """ valign=""top"" colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		'"                <p><b>Important Information from Snitz Forums 2000:</b>" & vbNewLine & _
		'"                <script type=""text/javascript"" src=""http://forum.snitz.com/forum/syndicate.asp""></script></p></font></td>" & vbNewLine & _
		'"              </tr>" & vbNewLine
Response.Write	"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine
WriteFooter
Response.End
%>