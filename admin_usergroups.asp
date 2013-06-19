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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	strQS = request.querystring
	Response.Redirect "admin_login.asp?target=" & Server.URLEncode(scriptname(ubound(scriptname)) & "?" & strQS)
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;<a href=""admin_usergroups.asp"">UserGroups&nbsp;Manager</a></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

UserGroupActionMode = LCase(Request.Querystring("mode"))
if UserGroupActionMode = "config" then
	for each key in Request.Form 
		if left(key,3) = "str" then
			strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLstring"))
		end if
	next
	
	Application(strCookieURL & "ConfigLoaded") = ""
	
	Call OkMessage("Configuration Posted","admin_usergroups.asp","Back To UserGroups Manager")
	
	WriteFooter
	Response.End
end if

If UserGroupActionMode = "delete" Or UserGroupActionMode = "modify" Then
	GroupID = Request.Querystring("ID")
End If

If UserGroupActionMode = "modify" Then
	strSql = "SELECT USERGROUP_NAME, USERGROUP_DESC, MEM_HIDE, MOD_HIDE, AUTOJOIN "
	strSql = strSql & "FROM " & strTablePrefix & "USERGROUPS "
	strSql = strSql & "WHERE USERGROUP_ID = " & GroupID
	set rsGroup = Server.CreateObject("ADODB.Recordset")
'	rsGroup.cachesize=20	'fix for mysql problem
	rsGroup.open strSql, my_Conn, 3
	GroupName = rsGroup("USERGROUP_NAME")
	GroupDesc = rsGroup("USERGROUP_DESC")
	MemHideGroup = rsGroup("MEM_HIDE")
	ModHideGroup = rsGroup("MOD_HIDE")
	GroupAutoJoin = rsGroup("AUTOJOIN")
	rsGroup.Close
	set rsGroup = Nothing
End If

Select Case UserGroupActionMode
	Case "delete"
		'## first delete usergroup from ALLOWED_USERGROUPS
		strSql = "DELETE FROM " & strTablePrefix & "ALLOWED_USERGROUPS "
		strSql = strSql & "WHERE USERGROUP_ID = " & GroupID
		my_Conn.Execute strSql
		'## next delete usergroup from USERGROUP_USERS where usergroup is a member of another usergroup
		strSql = "DELETE FROM " & strTablePrefix & "USERGROUP_USERS "
		strSql = strSql & "WHERE USERGROUP_ID = " & GroupID & " AND MEMBER_TYPE = 2"
		my_Conn.Execute strSql
		'## next delete usergroup from USERGROUP_USERS
		strSql = "DELETE FROM " & strTablePrefix & "USERGROUP_USERS "
		strSql = strSql & "WHERE USERGROUP_ID = " & GroupID
		'## last delete usergroup from USERGROUPS
		strSql = "DELETE FROM " & strTablePrefix & "USERGROUPS "
		strSql = strSql & "WHERE USERGROUP_ID = " & GroupID
		my_Conn.Execute strSql
		Response.Redirect("admin_usergroups.asp")
	Case "add", "modify"
		strAddOrModify = "Add Group"
		strConfirm = Request.Form("submit")
		If strConfirm = "Add Group" Then
			intAJ = 0
			intModH = 0
			intMemH = 0
			if Request.Form("GroupAutoJoin") = "on" then intAJ = 1
			if Request.Form("MemHideGroup") = "on" then intMemH = 1
			if Request.Form("ModHideGroup") = "on" then intModH = 1
			my_Conn.execute("INSERT INTO " & strTablePrefix & "USERGROUPS (USERGROUP_NAME,USERGROUP_DESC,AUTOJOIN,MEM_HIDE,MOD_HIDE) " & _
				"VALUES ('" & chkString(Request.Form("GroupName"),"SQLString") & "', '" & chkString(Request.Form("GroupDesc"),"SQLString") & "', " & intAJ & ", " & intMemH & ", " & intModH & ")")
			set rsGroup = Server.CreateObject("ADODB.Recordset")
			set rsGroup = my_conn.execute(TopSQL("SELECT USERGROUP_ID FROM " & strTablePrefix & "USERGROUPS ORDER BY USERGROUP_ID DESC;",1))
			Groupid = rsGroup("USERGROUP_ID")
			rsGroup.Close
			set rsGroup = nothing
			Response.Redirect("admin_usergroups.asp")
		End If
		If strConfirm = "Modify Group" Then
			'## check for usergroup name uniqueness
			intAJ = 0
			intModH = 0
			intMemH = 0
			if Request.Form("GroupAutoJoin") = "on" then intAJ = 1
			if Request.Form("MemHideGroup") = "on" then intMemH = 1
			if Request.Form("ModHideGroup") = "on" then intModH = 1
			strSql = "UPDATE " & strTablePrefix & "USERGROUPS SET USERGROUP_NAME = '" &_
				ChkString(Request.Form("GroupName"),"SQLString") & "', USERGROUP_DESC = '" &_
				ChkString(Request.Form("GroupDesc"),"SQLString") & "', AUTOJOIN = " &_
				intAJ & ", MEM_HIDE = " & intMemH & ", MOD_HIDE = " & intModH &_
				" WHERE USERGROUP_ID = " & cLng(GroupID)
			my_Conn.execute(strSql)
			Users  = Split(Request.Form("AuthUsers"),",")
			my_Conn.execute("DELETE FROM " & strTablePrefix & "USERGROUP_USERS WHERE USERGROUP_ID = " & cLng(GroupID))
			for count = Lbound(Users) to Ubound(Users)
				if Left(Trim(Users(count)),1) = "G" then
	  				intMemberType = 2
	  			else
	  				intMemberType = 1
	  			end if
	  			intUser = Trim(Users(count))
	  			intLen = Len(intUser)
	  			intUser = Mid(intUser, 2, intLen-1)
				strSql = "INSERT INTO " & strTablePrefix & "USERGROUP_USERS ("
				strSql = strSql & "USERGROUP_ID, MEMBER_ID, MEMBER_TYPE) VALUES ("
				strSql = strSql & GroupID & ", " & intUser & ", " & intMemberType & ")"
				my_Conn.execute(strSql)
			next

			my_Conn.execute("DELETE FROM " & strTablePrefix & "ALLOWED_USERGROUPS WHERE USERGROUP_ID = " & GroupID)
			for each key in Request.Form
				if Left(key,5) = "Perms" and Request.Form(key) <> "notset" then
					strFid = Trim(Mid(key,6))
					strSql = "INSERT INTO " & strTablePrefix & "ALLOWED_USERGROUPS ("
					strSql = strSql & "FORUM_ID, USERGROUP_ID, PERMS) VALUES ("
					strSql = strSql & strFid & ", " & GroupID & ", " & Request.Form(key) & ")"
					my_Conn.execute(strSql)
				end if
			next
			Response.Redirect("admin_usergroups.asp")
		End If
		Response.Write	"<form name=""GroupModify"" method=""POST"" action=""admin_usergroups.asp?mode=" & UserGroupActionMode & "&ID=" & GroupID & """>" & vbNewLine & _
						"<table class=""admin"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td colspan=""2"">Add/Modify Group</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Group Name:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<input type=""textbox"" name=""GroupName"" size=""52"" value=""" & ChkString(GroupName,"edit") & """ /></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Description:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<textarea name=""GroupDesc"" cols=""50"" rows=""3"">" & ChkString(GroupDesc,"edit") & "</textarea></td>" & vbNewLine & _
						"</tr>" & vbNewLine &_
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Group Options:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<input type=""checkbox"" name=""MemHideGroup"" value=""on""" & chkCheckBox(MemHideGroup,1,true) & ">&nbsp;Hidden from members<br />" & vbNewline &_
						"<input type=""checkbox"" name=""ModHideGroup"" value=""on""" & chkCheckBox(ModHideGroup,1,true) & ">&nbsp;Hidden from forum moderators<br />" & vbNewline &_
						"<input type=""checkbox"" name=""GroupAutoJoin"" value=""on""" & chkCheckBox(GroupAutoJoin,1,true) & ">&nbsp;Auto-join for new members on registration<br />" & vbNewline
		'if CInt(strEmail) = 1 and (CInt(strSubscription) >= 1 or CInt(strSubscription) <= 3) then response.write "<input type=""checkbox"" name=""GroupEnableSub"" value=""on"" onClick=""return false;"" disable>&nbsp;Enable forum subscriptions for this group (not available)<br />" & vbNewline
		Response.Write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine

		if UserGroupActionMode = "modify" then
			'## Group Member List
			Response.Write	"<SCRIPT LANGUAGE=""JavaScript"" SRC=""selectbox.js""></SCRIPT>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Group Members:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine

			strSql = "SELECT MEMBER_ID, M_NAME "
			strSql = strSql & "FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & "ORDER BY M_NAME"

			strSqlG = "SELECT USERGROUP_ID, USERGROUP_NAME "
			strSqlG = strSqlG & " FROM " & strTablePrefix & "USERGROUPS "
			strSqlG = strSqlG & " WHERE USERGROUP_ID <> " & GroupID

			'## find usergroups that include this usergroup as a member
			blnFirst = True
			ParentGroupList (GroupID)
			if blnFirst = "False" then strSqlG = strSqlG & ")"

			strSqlG = strSqlG & " ORDER BY USERGROUP_NAME"
			
			on error resume next
		
			set rsMember = my_Conn.execute (strSql)
			set rsGroup = my_Conn.execute(strSqlG)
			
			strSql = "SELECT MEMBER_ID FROM " & strTablePrefix & "USERGROUP_USERS "
			strSql = strSql & " WHERE USERGROUP_ID = " & GroupID
			strSql = strSql & " AND MEMBER_TYPE = 1"
			
			strSqlG = "SELECT MEMBER_ID FROM " & strTablePrefix & "USERGROUP_USERS "
			strSqlG = strSqlG & " WHERE USERGROUP_ID = " & GroupID
			strSqlG = strSqlG & " AND MEMBER_TYPE = 2"

			set rsGroupMember = my_Conn.execute (strSql)
			set rsSubGroup = my_Conn.execute(strSqlG)

			tmpStrUserList = ""
			do while not (rsGroupMember.EOF or rsGroupMember.BOF)
				if tmpStrUserList = "" then
					tmpStrUserList = rsGroupMember("MEMBER_ID")
				else
					tmpStrUserList = tmpStrUserList & "," & rsGroupMember("MEMBER_ID")
				end if
				rsGroupMember.movenext
			loop
			
			tmpStrGroupList = ""
			do while not (rsSubGroup.EOF or rsSubGroup.BOF)
				if tmpStrGroupList = "" then
					tmpStrGroupList = rsSubGroup("MEMBER_ID")
				else
					tmpStrGroupList = tmpStrGroupList & "," & rsSubGroup("MEMBER_ID")
				end if
				rsSubGroup.movenext
			loop
			SelectSize = 10
                  
			Response.Write	"<table width=""100%"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td style=""border:none;"" width=""50%"">" & vbNewLine & _
							"Forum Members:<br />" & vbNewLine & _
							"<select style=""width:100%"" name=""AuthUsersCombo"" size=""" & SelectSize & """ multiple onDblClick=""moveSelectedOptions(document.GroupModify.AuthUsersCombo, document.GroupModify.AuthUsers, false, '');sortSelect(document.GroupModify.AuthUsers)"">" & vbNewLine

			'## Pick from list
			rsGroup.movefirst
			do until rsGroup.eof
				if not(InStr("," & tmpStrGroupList & "," , "," & rsGroup("USERGROUP_ID") & ",") > 0) THEN
					Response.Write "<option value=""G" & rsGroup("USERGROUP_ID") & """" & isSel & ">* " & ChkString(rsGroup("USERGROUP_NAME"),"display") & "</option>" & vbNewline
				end if
				rsGroup.movenext
			loop
			set rsGroup = nothing
			rsMember.movefirst
			do until rsMember.eof
				if not(Instr("," & tmpStrUserList & "," , "," & rsMember("MEMBER_ID") & ",") > 0) then
					Response.Write "<option value=""M" & rsMember("MEMBER_ID") & """" & isSel & ">" & ChkString(rsMember("M_NAME"),"display") & "</option>" & vbNewline
				end if
				rsMember.movenext
			loop
			set rsMember = nothing			

			Response.Write	"</select>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"<td style=""border:none;"" width=""15""><br />" & vbNewLine & _
							"<a href=""javascript:moveAllOptions(document.GroupModify.AuthUsers, document.GroupModify.AuthUsersCombo, false, '');sortSelect(document.GroupModify.AuthUsersCombo)""><img src=""" & strImageURL & "icon_Private_remall.gif"" width=""23"" height=""22"" border=""0"" alt=""""></a>" & vbNewLine & _
							"<a href=""javascript:moveSelectedOptions(document.GroupModify.AuthUsers, document.GroupModify.AuthUsersCombo, false, '');sortSelect(document.GroupModify.AuthUsersCombo)""><img src=""" & strImageURL & "icon_Private_remove.gif"" width=""23"" height=""22"" border=""0"" alt=""""></a>" & vbNewLine & _
							"<a href=""javascript:moveSelectedOptions(document.GroupModify.AuthUsersCombo, document.GroupModify.AuthUsers, false, '');sortSelect(document.GroupModify.AuthUsers)""><img src=""" & strImageURL & "icon_Private_add.gif"" width=""23"" height=""22"" border=""0"" alt=""""></a>" & vbNewLine & _
							"<a href=""javascript:moveAllOptions(document.GroupModify.AuthUsersCombo, document.GroupModify.AuthUsers, false, '');sortSelect(document.GroupModify.AuthUsers)""><img src=""" & strImageURL & "icon_Private_addall.gif"" width=""23"" height=""22"" border=""0"" alt=""""></a>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"<td style=""border:none;"" width=""50%"">" & vbNewLine & _
							"Selected Members:<br>" & vbNewLine & _
							"<select style=""width:100%;"" name=""AuthUsers"" size=""" & SelectSize & """ multiple onDblClick=""moveSelectedOptions(document.GroupModify.AuthUsers, document.GroupModify.AuthUsersCombo, false, '');sortSelect(document.GroupModify.AuthUsersCombo)"">" & vbNewLine

			'## Selected List
			rsSubGroup.movefirst
			do until rsSubGroup.EOF
				if rsSubGroup("MEMBER_ID") <> "" then
					Response.Write "<option value=""G" & rsSubGroup("MEMBER_ID") & """>* " & ChkString(getUserGroupName(rsSubGroup("MEMBER_ID")),"display") & "</option>" & vbNewline
				end if
				rsSubGroup.movenext
			loop
			set rsSubGroup = nothing
			rsGroupMember.movefirst
			do until rsGroupMember.EOF
				if rsGroupMember("MEMBER_ID") <> "" then
					Response.Write "<option value=""M" & rsGroupMember("MEMBER_ID") & """>" & ChkString(getMemberName(rsGroupMember("MEMBER_ID")),"display") & "</option>" & vbNewline
				end if
				rsGroupMember.movenext
			loop
			set rsGroupMember = nothing

			Response.Write	"</select>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
			'## End Group Member List

			'## Begin Allowed Forums List
			Response.Write	"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Set Permissions on Forum(s):&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine &_
							"<table width=""100%"">" & vbNewLine &_
							"<tr class=""header"">" & vbNewLine & _
							"<td colspan=""2"">Forum</td>" & vbNewLine & _
							"<td>Permissions</td>" & vbNewLine & _
							"</tr>" & vbNewLine

			strSql = "SELECT CAT_ID, CAT_STATUS, CAT_NAME, CAT_ORDER "
			strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
			strSql = strSql & " ORDER BY CAT_ORDER, CAT_NAME "
			set rsCat = my_Conn.execute(strSql)
			arCat = Null
			if not rsCat.bof and not rsCat.eof then arCat = rsCat.GetRows
			rsCat.Close
			set rsCat = Nothing
			
			if not IsNull(arCat) then
				for CatCount = LBound(arCat,2) to UBound(arCat,2)
					response.write	"<tr class=""section"">" & vbNewLine & _
									"<td colspan=""3"">" & chkString(arCat(2,CatCount),"display") & "</td>" & vbNewLine & _
									"</tr>" & vbNewline
									
					strSql = "SELECT FORUM_ID, F_PRIVATEFORUMS, F_SUBJECT, CAT_ID, F_TYPE, F_ORDER "
					strSql = strSql & "FROM " & strTablePrefix & "FORUM"
					strSql = strSql & " WHERE CAT_ID = " & arCat(0,CatCount)
					strSql = strSql & " ORDER BY F_ORDER ASC, F_SUBJECT ASC;"
					set rsF = my_Conn.execute(strSql)
					arF = Null
					if not rsF.bof and not rsF.eof then arF = rsF.GetRows
					rsF.Close
					set rsF = Nothing
					if not IsNull(arF) then
						for FCount = LBound(arF,2) to UBound(arF,2)
							strSql = "SELECT PERMS FROM " & strTablePrefix & "ALLOWED_USERGROUPS WHERE "
							strSql = strSql & "FORUM_ID = " & arF(0,FCount)
							strSql = strSql & " AND USERGROUP_ID = " & GroupID
							set rsPerms = my_Conn.execute(strSql)
							strPerms = ""
							if not rsPerms.BOF and not rsPerms.EOF then strPerms = rsPerms(0)
							rsPerms.Close
							set rsPerms = Nothing
							response.write	"<tr class=""content"">" & vbNewline &_
											"<td class=""options""><a href=""post.asp?method=EditForum&FORUM_ID=" & arF(0,FCount) & "&CAT_ID=" & arCat(0,CatCount) & "&type=0"">" & getCurrentIcon(strIconFolderPencil,"Edit Forum Properties","hspace=""0""") & "</a></td>" & vbNewline &_
											"<td>" & chkString(arF(2,FCount),"display") & "</td>" & vbNewline &_
											"<td class=""options""><select name=""Perms" & arF(0,FCount) & """>" & vbNewline &_
											"<option value=""notset"""
							if strPerms = "" then response.write " selected"
							response.write	">Do Not Set</option>" & vbNewline
							if (arF(1,FCount) = 1 or arF(1,FCount) = 6 or arF(1,FCount) = 3) then
								response.write	"<option value=""0"""
								if strPerms = "0" then response.write " selected"
								response.write	">Allow</option>" & vbNewline
							end if
							response.write	"<option value=""2"""
							if strPerms = "2" then response.write " selected"
							response.write	">Read-Only</option>" & vbNewline
							response.write	"<option value=""1"""
							if strPerms = "1" then response.write " selected"
							response.write	">Deny</option>" & vbNewline &_
											"</select></td>" & vbNewline &_
											"</tr>" & vbNewline
						next
					end if
				next
			else
				response.write	"<tr><td colspan=""3""><font size=""" & strDefaultFontSize & """>No forums were found.</font></td></tr>" & vbNewline
			end if

			response.write	"</table>" & vbNewline &_
							"</td>" & vbNewline &_
							"</tr>" & vbNewline
			'## End Allowed Forums List
	
			strAddOrModify = "Modify Group"
		end if 'end Modify mode

		Response.Write	"<tr class=""options"">" & vbNewLine & _
						"<td colspan=""2"">" & vbNewLine & _
						"<input type=""hidden"" name=""submit"" value=""" & strAddOrModify & """ />" & vbNewLine & _
						"<button type=""submit"""
		if UserGroupActionMode = "modify" then 
			Response.Write	" onClick=""selectAllOptions(document.GroupModify.AuthUsers)"""
		end if
		Response.Write	">" & strAddOrModify & "</button>" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</form>" & vbNewLine

	Case Else
		Response.Write	"<script type=""text/javascript"" language=""javascript"">" & vbNewline &_
						"		function limitValues() {" & vbNewline &_
						"			var UGView = document.getElementById('strUGView');" & vbNewline &_
						"			var UGForm = UGView.form;" & vbNewline &_
						"			switch (UGView.value) {" & vbNewline &_
						"			case '0':" & vbNewline &_
						"				UGForm.strUGMemView[1].disabled = true;" & vbNewline &_
						"				UGForm.strUGMemView[2].disabled = true;" & vbNewline &_
						"				UGForm.strUGMemView[0].checked = true;" & vbNewline &_
						"				break;" & vbNewline &_
						"			case '1':" & vbNewline &_
						"				UGForm.strUGMemView[1].disabled = false;" & vbNewline &_
						"				UGForm.strUGMemView[2].disabled = true;" & vbNewline &_
						"				if (UGForm.strUGMemView[2].checked == true) {UGForm.strUGMemView[1].checked = true;}" & vbNewline &_
						"				break;" & vbNewline &_
						"			case '2':" & vbNewline &_
						"				UGForm.strUGMemView[1].disabled = false;" & vbNewline &_
						"				UGForm.strUGMemView[2].disabled = false;" & vbNewline &_
						"				break;" & vbNewline &_
						"			default:" & vbNewline &_
						"				UGForm.strUGMemView[1].disabled = true;" & vbNewline &_
						"				UGForm.strUGMemView[2].disabled = true;" & vbNewline &_
						"				UGForm.strUGMemView[0].checked = true;" & vbNewline &_
						"				break;" & vbNewline &_
						"			}" & vbNewline &_
						"		}" & vbNewline &_
						"</script>" & vbNewline
		
		Response.Write	"<form action=""admin_usergroups.asp?mode=config"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
						"<table class=""admin"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td colspan=""2"">UserGroups Configuration</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">UserGroup Information Visible to:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<select name=""strUGView"" onChange=""limitValues();"">" & vbNewline &_
						"<option value=""0""" & chkSelect(strUGView,0) & ">Admin only</option>" & vbNewline &_
						"<option value=""1""" & chkSelect(strUGView,1) & ">UserGroup Members</option>" & vbNewline &_
						"<option value=""2""" & chkSelect(strUGView,2) & ">All Members</option>" & vbNewline &_
						"</select>" & vbNewline &_
						"</td>" & vbNewline &_
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">UserGroup Members Visible to:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<input type=""radio"" name=""strUGMemView"" value=""0""" & chkRadio(strUGMemView,0,true) & " />Admin only&nbsp;" & vbNewline &_
						"<input type=""radio"" name=""strUGMemView"" value=""1""" & chkRadio(strUGMemView,1,true) & " />UserGroup Members&nbsp;" & vbNewline &_
						"<input type=""radio"" name=""strUGMemView"" value=""2""" & chkRadio(strUGMemView,2,true) & " />All Members" & vbNewline &_
						"</td>" & vbNewline &_
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Moderator Access to Forum UserGroup Permissions:&nbsp;<br /><span class=""smtext"">(View and Edit allow moderators to view UserGroup members)</span>&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<select name=""strUGModForums"">" & vbNewline &_
						"<option value=""0""" & chkSelect(strUGModForums,0) & ">None</option>" & vbNewline &_
						"<option value=""1""" & chkSelect(strUGModForums,1) & ">View</option>" & vbNewline &_
						"<option value=""2""" & chkSelect(strUGModForums,2) & ">Edit</option>" & vbNewline &_
						"</select>" & vbNewline &_
						"</td>" & vbNewline &_
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""options"" colspan=""2""><button type=""submit"" id=""submit1"" name=""submit1"">Submit New Config</button></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</form><br />" & vbNewLine &_
						"<script type=""text/javascript"" language=""javascript"">limitValues();</script>" & vbNewline
		
		Response.Write	"<table class=""admin"" width=""100%"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td>Group Name</td>" & vbNewLine & _
						"<td>Description</td>" & vbNewLine & _
						"<td class=""options"">" & _
						"<a href=""admin_usergroups.asp?mode=Add"">" & getCurrentIcon(strIconGroup,"Add UserGroup ...","hspace=""0""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine

		'## Forum_SQL - Find all records with the search criteria in them
		strSql = "SELECT " & strTablePrefix & "USERGROUPS.USERGROUP_ID, "
		strSql = strSql & strTablePrefix & "USERGROUPS.USERGROUP_NAME, " 
		strSql = strSql & strTablePrefix & "USERGROUPS.USERGROUP_DESC "
		strSql = strSql & " FROM " &strTablePrefix & "USERGROUPS "

		set rs = Server.CreateObject("ADODB.Recordset")
		rs.cachesize=20
		rs.open  strSql, my_Conn, 3
		if rs.EOF or rs.BOF then  '## No Groups found
			Response.Write	"<tr>" & vbNewLine & _
							"<td colspan=""3"">No Groups Found.</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		else
			do until rs.EOF
				Response.Write	"<tr>" & vbNewLine & _
								"<td>" & ChkString(rs("USERGROUP_NAME"),"display") & "</td>" & vbNewLine & _
								"<td>" & ChkString(rs("USERGROUP_DESC"),"display") & "</td>" & vbNewLine & _
								"<td class=""options"">" & vbNewLine & _
								"<a href=""admin_usergroups.asp?mode=Modify&ID=" & rs("USERGROUP_ID") & """>" & getCurrentIcon(strIconPencil,"Modify UserGroup", "hspace=""0""") & "</a>" & vbNewLine & _
								"<a href=""admin_usergroups.asp?mode=Delete&ID=" & rs("USERGROUP_ID") & """>" & getCurrentIcon(strIconTrashcan,"Delete UserGroup", "hspace=""0""") & "</a>" & vbNewLine & _
								"</td>" & vbNewLine & _
								"</tr>" & vbNewLine
				rs.MoveNext
			loop
		end if 
		rs.Close
		set rs = Nothing
		Response.Write	"</table>" & vbNewLine

	End Select

WriteFooter
Response.End

sub ParentGroupList(intGroupID)
	strSqlParG = "SELECT USERGROUP_ID FROM " & strTablePrefix & "USERGROUP_USERS"
	strSqlParG = strSqlParG & " WHERE MEMBER_ID = " & intGroupID & " AND MEMBER_TYPE = 2"
	set rsParent = my_Conn.execute(strSqlParG)
	if not rsParent.bof and not rsParent.eof then
		myGroup = rsParent.GetRows()
		set rsParent = Nothing
		for i = LBound(myGroup) to UBound(myGroup)
			if blnFirst = False Then
				strSqlG = strSqlG & (", ")
			else
				blnFirst = False
				strSqlG = strSqlG & " AND USERGROUP_ID NOT IN ("
			end if
			strSqlG = strSqlG & myGroup(0,i)
			ParentGroupList(myGroup(0,i))
		next
	end if
end sub
%>
