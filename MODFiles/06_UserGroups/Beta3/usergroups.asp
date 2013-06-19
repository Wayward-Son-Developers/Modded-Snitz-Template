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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<%
UG_Err_Msg = ""

mode = Request.QueryString("mode")
if mode = "" then mode = "ViewGroups"
if mode = "ViewUsers" then
	GroupID = Request.QueryString("ID")
	if not IsNumeric(GroupID) then response.redirect("usergroups.asp")
	if not InStr("," & Trim(Session(strCookieURL & "UserGroups" & MemberID)) & ",", "," & Trim(GroupID) & ",") > 0 then
		if mlev < 3 or (mlev = 3 and CInt(strUGModForums) = 0) then
			if CInt(strUGMemView) < 2 then UG_Err_Msg = "You must be a Member of this UserGroup to view its Membership."
		end if
	end if
	strSql = "SELECT USERGROUP_NAME " &_
		"FROM " & strTablePrefix & "USERGROUPS " &_
		"WHERE USERGROUP_ID = " & GroupID
	set rsGroup = my_Conn.execute(strSql)
	if not rsGroup.bof and not rsGroup.eof then
		strGroupName = rsGroup("USERGROUP_NAME")
	else
		response.redirect("usergroups.asp")
	end if
	rsGroup.close
	set rsGroup = Nothing
end if

blnCanView = chkUserGroupView(MemberID)

Response.Write	"	<table width=""100%"" border=""0"">" & vbNewLine & _
		"		<tr>" & vbNewLine & _
		"			<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"				" & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
		"				" & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconGroup,"","") & "&nbsp;<a href=""usergroups.asp"">UserGroup Information</a><br /> " & vbNewLine
if mode = "ViewUsers" then Response.write "				" & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconGroup,"","") & "&nbsp;" & strGroupName & "<br /> " & vbNewLine
Response.Write	"			</font></td>" & vbNewLine & _
		"		</tr>" & vbNewLine & _
		"	</table>" & vbNewLine &_
		"	<br />" & vbNewLine

if (blnCanView = False or UG_Err_Msg <> "") then
	if UG_Err_Msg = "" then UG_Err_Msg = "You do not have access to UserGroups."
	Response.Write	"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem!</font></p>" & vbNewLine & _
			"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>" & UG_Err_Msg & "</font></p>" & vbNewLine & _
			"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Back to Forum</a></font></p>" & vbNewLine & _
			"	<br />" & vbNewLine
	WriteFooter
	Response.End
end if



Response.Write	"	<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"		<tr>" & vbNewLine & _
		"			<td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"				<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">" & vbNewLine

Select Case mode
	Case "ViewGroups"
		'## grab UserGroups from the db
		strSql = "SELECT USERGROUP_ID, USERGROUP_NAME, USERGROUP_DESC, MEM_HIDE " &_
			"FROM " & strTablePrefix & "USERGROUPS"
		Select Case mlev
			Case 4	'## do nothing
			Case 3
				if CInt(strUGModForums) > 0 then
					strSql = strSql & " WHERE MOD_HIDE = 0"
				else
					strSql = strSql & " WHERE MEM_HIDE = 0"
					if CInt(strUGView) = 1 then strSql = strSql & " AND USERGROUP_ID IN (" & Trim(Session(strCookieURL & "UserGroups" & MemberID)) & ")"
				end if
			Case Else
				strSql = strSql & " WHERE MEM_HIDE = 0"
				if CInt(strUGView) = 1 then strSql = strSql & " AND USERGROUP_ID IN (" & Trim(Session(strCookieURL & "UserGroups" & MemberID)) & ")"
		End Select
		strSql = strSql & " ORDER BY USERGROUP_NAME"
		set rsGroups = my_Conn.execute(strSql)
		arGroups = Null
		if not rsGroups.bof and not rsGroups.eof then arGroups = rsGroups.GetRows
		rsGroups.close
		set rsGroups = Nothing
		
		Response.Write	"					<tr>" & vbNewLine &_
				"						<td bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>UserGroup Name</font></b></td>" & vbNewLine & _
				"						<td bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>UserGroup Description</font></b></td>" & vbNewLine & _
				"						<td nowrap align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>"
		if mlev = 4 then
			response.write	"<a href=""admin_usergroups.asp"" alt=""UserGroups Manager"">" & getCurrentIcon(strIconGroup,"UserGroups Manager","hspace=""0""") & "</a>"
		else
			response.write	"&nbsp;"
		end if
		Response.write	"</font></b></td>" & vbNewLine & _
				"					</tr>" & vbNewLine

		if not IsNull(arGroups) then
			for GCnt = LBound(arGroups,2) to UBound(arGroups,2)
				Response.Write	"					<tr>" & vbNewLine &_
						"						<td valign=""top"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & ChkString(arGroups(1,GCnt),"display") & "</font></td>" & vbNewLine & _
						"						<td valign=""top"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & ChkString(arGroups(2,GCnt),"display") & "</font></td>" & vbNewLine & _
						"						<td nowrap valign=""top"" align=""center"" bgcolor=""" & strForumCellColor & """>" & vbNewline &_
						"							<a href=""usergroups.asp?mode=ViewUsers&ID=" & arGroups(0,GCnt) & """ alt=""View Users in this UserGroup"">" & getCurrentIcon(strIconGroup,"View Users in this UserGroup","hspace=""0""") & "</a>" & vbNewline
				if mlev = 4 then response.write "							<a href=""admin_usergroups.asp?mode=Modify&ID=" & arGroups(0,GCnt) & """ alt=""Modify this UserGroup"">" & getCurrentIcon(strIconPencil,"Modify this UserGroup","hspace=""0""") & "</a>" & vbNewline
				Response.write	"						</td>" & vbNewLine & _
						"					</tr>" & vbNewLine
			next
		else
			Response.Write	"					<tr><td colspan=""3"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>No usergroups were found.</font></td></tr>" & vbNewLine
		end if
	Case "ViewUsers"
		'## grab Users in this UserGroup
		strSql = "SELECT M.MEMBER_ID, M.M_NAME  FROM " & strMemberTablePrefix & "MEMBERS M " &_
			"INNER JOIN " & strTablePrefix & "USERGROUP_USERS UM ON M.MEMBER_ID = UM.MEMBER_ID " &_
			"WHERE UM.MEMBER_TYPE = 1 AND UM.USERGROUP_ID = " & GroupID & " ORDER BY M.M_NAME"
		set rsUGMem = my_Conn.execute(strSql)
		arUGMem1 = Null
		if not rsUGMem.bof and not rsUGMem.eof then arUGMem1 = rsUGMem.GetRows
		rsUGMem.close
		set rsUGMem = Nothing
		strSql = "SELECT U.USERGROUP_ID, U.USERGROUP_NAME FROM " & strTablePrefix & "USERGROUPS U " &_
			"INNER JOIN " & strTablePrefix & "USERGROUP_USERS UM ON U.USERGROUP_ID = UM.MEMBER_ID " &_
			"WHERE UM.MEMBER_TYPE = 2 AND UM.USERGROUP_ID = " & GroupID & " ORDER BY U.USERGROUP_NAME"
		set rsUGMem = my_Conn.execute(strSql)
		arUGMem2 = Null
		if not rsUGMem.bof and not rsUGMem.eof then arUGMem2 = rsUGMem.GetRows
		rsUGMem.close
		set rsUGMem = Nothing
		if IsNull(arUGMem1) and IsNull(arUGMem2) then blnNoMem = true
		Response.Write	"					<tr>" & vbNewLine &_
				"						<td bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Member Name</font></b></td>" & vbNewLine & _
				"						<td bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Member Type</font></b></td>" & vbNewLine & _
				"						<td nowrap align=""center"" bgcolor=""" & strHeadCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&nbsp;</font></td>" & vbNewLine & _
				"					</tr>" & vbNewLine
		if not IsNull(arUGMem2) then
			for UCnt = LBound(arUGMem2,2) to UBound(arUGMem2,2)
				Response.Write	"					<tr>" & vbNewLine &_
						"						<td valign=""top"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & ChkString(arUGMem2(1,UCnt),"display") & "</font></td>" & vbNewLine & _
						"						<td valign=""top"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>UserGroup</font></td>" & vbNewLine & _
						"						<td nowrap valign=""top"" align=""center"" bgcolor=""" & strForumCellColor & """>" & vbNewline &_
						"							<a href=""usergroups.asp?mode=ViewUsers&ID=" & arUGMem2(0,UCnt) & """ alt=""View UserGroup Members"">" & getCurrentIcon(strIconGroup,"View UserGroup Members","hspace=""0""") & "</a>" & vbNewline
				if mlev = 4 then response.write "							<a href=""admin_usergroups.asp?mode=Modify&ID=" & arUGMem2(0,UCnt) & """ alt=""Modify this UserGroup"">" & getCurrentIcon(strIconPencil,"Modify this UserGroup","hspace=""0""") & "</a>" & vbNewline
				Response.write	"						</td>" & vbNewLine & _
						"					</tr>" & vbNewLine
			next
		end if
		if not IsNull(arUGMem1) then
			for UCnt = LBound(arUGMem1,2) to UBound(arUGMem1,2)
				Response.Write	"					<tr>" & vbNewLine &_
						"						<td valign=""top"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & ChkString(arUGMem1(1,UCnt),"display") & "</font></td>" & vbNewLine & _
						"						<td valign=""top"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>User</font></td>" & vbNewLine & _
						"						<td nowrap valign=""top"" align=""center"" bgcolor=""" & strForumCellColor & """>" & vbNewline &_
						"							<a href=""pop_profile.asp?mode=display&id=" & arUGMem1(0,UCnt) & """ alt=""View Member's Profile"">" & getCurrentIcon(strIconProfile,"View Member's Profile","hspace=""0""") & "</a>" & vbNewline
				if mlev = 4 then response.write "							<a href=""pop_profile.asp?mode=Modify&ID=" & arUGMem1(0,UCnt) & """ alt=""Modify this Member"">" & getCurrentIcon(strIconPencil,"Modify this Member","hspace=""0""") & "</a>" & vbNewline
				Response.write	"						</td>" & vbNewLine & _
						"					</tr>" & vbNewLine
			next
		end if
		if blnNoMem = true then
			Response.Write	"					<tr><td colspan=""3"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>This UserGroup does not have any members.</font></td></tr>" & vbNewLine
		end if
End Select

Response.Write	"				</table>" & vbNewline &_	
		"			</td>" & vbNewline &_
		"		</tr>" & vbNewline &_
		"	</table>" & vbNewline &_
		"    <br />" & vbNewLine

WriteFooter
Response.End
%>
