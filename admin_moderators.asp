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
If IsNumeric(Request.QueryString("Forum")) Then
	Forum_ID = cLng(Request.QueryString("Forum"))
Else
	Forum_ID = 0
End If
If IsNumeric(Request.QueryString("userid")) Then
	User_ID = cLng(Request.QueryString("userid"))
Else
	User_ID = 0
End If
If IsNumeric(Request.QueryString("action")) Then
	Action_ID = cInt(Request.QueryString("action"))
Else
	Action_ID = 0
End If
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp"-->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Moderator Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

Response.Write	"<table class=""admin"" width=""75%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td>Moderator Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine

if Forum_ID = 0 then
	Response.Write	"<tr>" & vbNewLine & _
					"<td>" & vbNewLine
	Response.Write	"<p>Select a forum to edit moderators for that forum:</p><ul>"
	'## Forum_SQL
	strSql = "SELECT C.CAT_ORDER, C.CAT_NAME, F.CAT_ID, F.FORUM_ID, F.F_ORDER, F.F_SUBJECT " &_
	" FROM " & strTablePrefix & "CATEGORY C, " & strTablePrefix & "FORUM F" &_
	" WHERE C.CAT_ID = F.CAT_ID "
	strSql = strSql & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT ASC;"

	set rs = my_Conn.Execute(strSql)

	if rs.eof or rs.bof then
		'nothing
	else
		iOldCat = rs("CAT_ID")
		do until rs.EOF
			Response.Write	"<li>" & rs("CAT_NAME") & "<ul>" & vbNewLine
			iNewCat = rs("CAT_ID")
			if iNewCat <> iOldCat Then
				Response.Write	"</ul></li><li>" & rs("CAT_NAME") & "<ul>" & vbNewLine
				iOldCat = iNewCat
			end if
			Response.Write	"<li><a href=""admin_moderators.asp?forum=" & rs("FORUM_ID") & """>" & rs("F_SUBJECT") & "</a></li>" & vbNewLine
			rs.MoveNext
		loop
	end if
else
	if Action_ID = 0 then
		if User_ID = 0 then
			Response.Write	"<tr>" & vbNewLine & _
							"<td>" & vbNewLine
			Response.Write	"<p>Select a user to grant/revoke moderator powers for that user:<br />(Users in bold are currently moderators of this forum.)</p>"
			'## Forum_SQL
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_LEVEL > 1 "
			strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
			strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"

			set rs = my_Conn.Execute(strSql)

			Response.Write	"<ul>" & vbNewLine
			do until rs.EOF
				Response.Write	"<li>"
				if chkForumModerator(Forum_ID, rs("M_NAME")) then Response.Write("<b>")
				Response.Write	"<a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & rs("MEMBER_ID")& """>" & rs("M_NAME") & "</a>"
				If chkForumModerator(Forum_ID, rs("M_NAME")) then Response.Write("</b>")
				Response.Write	"</li>" & vbNewLine
				rs.MoveNext
			loop
			Response.Write	"</ul>" & vbNewLine
		else
			Response.Write	"<tr class=""warning"">" & vbNewLine & _
							"<td>" & vbNewLine
			'## Forum_SQL
			strSql = "SELECT " & strTablePrefix & "MODERATOR.FORUM_ID, " & strTablePrefix & "MODERATOR.MEMBER_ID, " & strTablePrefix & "MODERATOR.MOD_TYPE "
			strSql = strSql & " FROM " & strTablePrefix & "MODERATOR "
			strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & User_ID & " "
			strSql = strSql & " AND " & strTablePrefix & "MODERATOR.FORUM_ID = " & Forum_ID & " "

			set rs = my_Conn.Execute(strSql)

			if rs.EOF then
				Response.Write	"<p>The selected user is not a moderator of the selected forum.</p>" & vbNewLine & _
								"<p>If you would like to make this user the moderator of this forum, " & _
								"<a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & User_ID & "&action=1"">click here</a>.</p>" & vbNewLine
			else
				Response.Write	"<p>The selected user is currently a moderator of the selected forum.</p>" & vbNewLine & _
								"<p>If you would like to remove this user's moderator status in this forum, " & _
								"<a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & User_ID & "&action=2"">click here</a>.</p>" & vbNewLine
			end if
		end if
	else
		Response.Write	"<tr class=""ok"">" & vbNewLine & _
						"<td>" & vbNewLine
		select case Action_ID
			case 1
				'## Forum_SQL
				strSql = "INSERT INTO " & strTablePrefix & "MODERATOR "
				strSql = strSql & "(FORUM_ID"
				strSql = strSql & ", MEMBER_ID"
				strSql = strSql & ") VALUES (" 
				strSql = strSql & Forum_ID
				strSql = strSql & ", " & User_ID
				strSql = strSql & ")"

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"<p>The selected user is now a moderator of the selected forum.</p>" & vbNewLine & _
								"<p><a href=""admin_moderators.asp"">Back to Moderator Options</a></p>" & vbNewLine
			case 2

				'## Forum_SQL
				strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
				strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & Forum_ID & " "
				strSql = strSql & " AND   " & strTablePrefix & "MODERATOR.MEMBER_ID = " & User_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"<p>The selected user's moderator status in the selected forum has been removed.</p>" & vbNewLine & _
								"<p><a href=""admin_moderators.asp"">Back to Moderator Options</a></p>" & vbNewLine
		end select
	end if
end if
Response.Write	"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

WriteFooter
Response.End
%>
