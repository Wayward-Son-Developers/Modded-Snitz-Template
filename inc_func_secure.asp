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


function getMemberName(fUser_Number)

	dim strSql
	dim rsGetmemberName

	'## Forum_SQL
	if isNull(fUser_Number) then exit function
	strSql = "SELECT M_NAME "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE MEMBER_ID = " & cLng(fUser_Number)

	set rsGetMemberName = Server.CreateObject("ADODB.Recordset")
	rsGetMemberName.open strSql, my_Conn

	if rsGetMemberName.EOF or rsGetMemberName.BOF then
		getMemberName = ""
	else
		getMemberName = chkString(rsGetMemberName("M_NAME"),"display")
	end if

	rsGetMemberName.close
	set rsGetMemberName = nothing

end function


function getMemberID(fUser_Name)

	dim strSql
	dim rsGetMemberID

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(fUser_Name, "SQLString") & "'"

	set rsGetMemberID = Server.CreateObject("ADODB.Recordset")
	rsGetMemberID.open strSql, my_Conn

	if rsGetMemberID.EOF then
		getMemberID = 0
	else
		getMemberID = cLng(rsGetMemberID("MEMBER_ID"))
	end if

	rsGetMemberID.close
	set rsGetMemberID = nothing

end function


function chkDisplayForum(fPrivateForums,fFPasswordNew,fForum_ID,UserNum)

	dim strSql
	dim rsAccess
	chkDisplayForum = false

	if (mLev = 4) or (mLev = 3 and ModerateAllowed = "Y") then
		chkDisplayForum = true
		exit function
	end if
	
	if isDeniedMember(fForum_ID, UserNum) = 1 then
		chkDisplayForum = false
		exit function
	end if
	
	select case cLng(fPrivateForums)
		case 0, 1, 2, 3, 4, 7, 9
			chkDisplayForum = true
			exit function
		case 5
			if UserNum = -1 then
				chkDisplayForum = false
				exit function
			else
				chkDisplayForum = true
				exit function
			end if
		case 6
			if UserNum = -1 then
				chkDisplayForum = false
				exit function
			end if
			if isAllowedMember(fForum_ID,UserNum) = 1 then
				chkDisplayForum = true
			else
				chkDisplayForum = false
			end if 
	 	case 8
			chkDisplayForum = false
			if strAuthType ="nt" THEN
				NTGroupSTR = Split(Session(strCookieURL & "strNTGroupsSTR"), ", ")
				for j = 0 to ubound(NTGroupSTR)
					NTGroupDBSTR = Split(fFPasswordNew, ", ")
					for i = 0 to ubound(NTGroupDBSTR)
						if NTGroupDBSTR(i) = NTGroupSTR(j) then
							chkDisplayForum = true
							exit function
						end if
					next
				next
			end if
		case else 
			chkDisplayForum = true
	end select 
end function


function chkForumAccess(fForum, UserNum, Display)
	if MemberID = UserNum then
		if mLev < 1 then
			chkForumAccess = false
		elseif mLev = 3 then
			chkForumAccess = true
		elseif mLev = 4 then
			chkForumAccess = true
			exit function
		end if
	end if

	'## Forum_SQL
	strSql = "SELECT F_PRIVATEFORUMS, F_SUBJECT, F_PASSWORD_NEW "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE FORUM_ID = " & cLng(fForum)

	Set rsStatus = Server.CreateObject("ADODB.Recordset")
	rsStatus.open strSql, my_Conn

	if rsStatus.EOF or rsStatus.BOF then
		rsStatus.close
		set rsStatus = nothing
		Response.Redirect("default.asp")
	else
		if isDeniedMember(fForum, UserNum) = 1 then
			chkForumAccess = false
			if Display then
				doNotAllowed
				Response.End
			end if
			exit function
		end if
		
		dim Users
		dim MatchFound
		If rsStatus("F_PRIVATEFORUMS") <> 0 then

			Select case rsStatus("F_PRIVATEFORUMS")
				case 0
					chkForumAccess = true
				case 1, 6 '## Allowed Users
					if isAllowedMember(fForum,UserNum) = 1 then
						chkForumAccess = true
					else
						if Display then
							doNotAllowed
							Response.end
						else
							chkForumAccess = false
						end if
					end if
				case 2 '## password
					select case Request.Cookies(strUniqueID & "Forum")("PRIVATE_" & rsStatus("F_SUBJECT"))
						case rsStatus("F_PASSWORD_NEW")
							chkForumAccess = true
						case else
							if Request("pass") = "" then
								if Display then
									doPasswordForm
									Response.End
								else								
									chkForumAccess = false
								end if
							else
								if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
									if Display then
										Call FailMessage("<li>Invalid password!</li>",True)
										WriteFooter
										Response.End
									else
										chkForumAccess = false
									end if
								else
									if strSetCookieToForum = 1 then
										Response.Cookies(strUniqueID & "Forum").Path = strCookieURL
									end if
									Response.Cookies(strUniqueID & "Forum")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
									chkForumAccess = true
								end if
							end if
					end select
				case 3 '## Either Password or Allowed
					if isAllowedMember(fForum,UserNum) = 1 then
						chkForumAccess = true
					else
						chkForumAccess = false
					end if
					if not(chkForumAccess) then 
						select case Request.Cookies(strUniqueID & "Forum")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								chkForumAccess = true
							case else
								if Request("pass") = "" then
									if Display then
										doPasswordForm
										Response.End
									else								
										chkForumAccess = false
									end if
								else
									if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
										if Display then
											Call FailMessage("<li>Invalid password!</li>",True)
											WriteFooter
											Response.End
										else
											chkForumAccess = false
										end if
									else
										if strSetCookieToForum = 1 then
											Response.Cookies(strUniqueID & "Forum").Path = strCookieURL
										end if
										Response.Cookies(strUniqueID & "Forum")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
										chkForumAccess = true
									end if
								end if
						end select
					end if
				'## code added 07/13/2000
				case 7 '## members or password
					if strDBNTUserName = "" then
						select case Request.Cookies(strUniqueID & "Forum")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								chkForumAccess = true
							case else
								if Request("pass") = "" then
									if Display then
										doLoginForm
										response.end
									else
										chkForumAccess = false
									end if
								else
									if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
										if Display then
											Call FailMessage("<li>Invalid password!</li>",True)
											WriteFooter
											Response.End
										else
											chkForumAccess = false
										end if
									else
										if strSetCookieToForum = 1 then
											Response.Cookies(strUniqueID & "Forum").Path = strCookieURL
										end if
										Response.Cookies(strUniqueID & "Forum")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
										chkForumAccess = true
									end if
								end if
						end select
					else
						chkForumAccess = true
					end if
					'## end code added 07/13/2000
					
				case 4, 5 '## members only
					if Usernum = -1 or Usernum = "" then
						if Display then
							doNotLoggedInForm
						else
							chkForumAccess = false
						end if
					else '## V3.1 SR4
						chkForumAccess = true
					end if

				case 8, 9
					test="test db"
					chkForumAccess = FALSE
					if strAuthType="db" then
						chkForumAccess = true
						rsStatus.close
						set rsStatus = nothing
						exit function
					end if
					NTGroupSTR = Split(Session(strCookieURL & "strNTGroupsSTR"), ", ")
					for j = 0 to ubound(NTGroupSTR)
						NTGroupDBSTR = Split(rsStatus("F_PASSWORD_NEW"), ", ")
						for i = 0 to ubound(NTGroupDBSTR)
							if NTGroupDBSTR(i) = NTGroupSTR(j) then
								chkForumAccess = True
								rsStatus.close
								set rsStatus = nothing
								exit function
							end if
						next
					next
					if Display then
						doNotAllowed
					end if
				case else
					chkForumAccess = true
			end select
		else
			chkForumAccess = true
		end if
	end if

	rsStatus.close
	set rsStatus = nothing
end function


function chkForumAccessNew(fPrivateForums,fFPasswordNew,fForum_Subject,fForum_ID,UserNum)
	if MemberID = UserNum then
		if mLev < 1 then
			chkForumAccessNew = false
		elseif mLev = 3 then
			chkForumAccessNew = true
		elseif mLev = 4 then
			chkForumAccessNew = true
			exit function
		end if
	end if
	
	if isDeniedMember(fForum_ID, UserNum) = 1 then
		chkForumAccessNew = false
		exit function
	end if
	
	dim Users
	dim MatchFound

	Select case fPrivateForums
		case 0
			chkForumAccessNew = true
		case 1, 6 '## Allowed Members List
			if isAllowedMember(fForum_ID,UserNum) = 1 then
				chkForumAccessNew = true
			else
				chkForumAccessNew = false
			end if
		case 2 '## password
			select case Request.Cookies(strUniqueID & "Forum")("PRIVATE_" & fForum_Subject)
				case fFPasswordNew
					chkForumAccessNew = true
				case else
					chkForumAccessNew = false
			end select
		case 3 '## Either Password or Allowed Members List
			if isAllowedMember(fForum_ID,UserNum) = 1 then
				chkForumAccessNew = true
			else
				chkForumAccessNew = false
			end if
			if not(chkForumAccessNew) then 
				select case Request.Cookies(strUniqueID & "Forum")("PRIVATE_" & fForum_Subject)
					case fFPasswordNew
						chkForumAccessNew = true
					case else
						chkForumAccessNew = false
				end select
			end if
		case 7 '## Members or Password
			if Usernum = -1 or Usernum = "" then
				select case Request.Cookies(strUniqueID & "Forum")("PRIVATE_" & fForum_Subject)
					case fFPasswordNew
						chkForumAccessNew = true
					case else
						chkForumAccessNew = false
				end select
			else
				chkForumAccessNew = true
			end if
			
		case 4, 5 '## Members only
			if Usernum = -1 or Usernum = "" then
				chkForumAccessNew = false
			else
				chkForumAccessNew = true
			end if
		case 8, 9 '## NT Global Groups
			test="test db"
			chkForumAccessNew = false
			if strAuthType="db" then
				chkForumAccessNew = true
			end if
			NTGroupSTR = Split(Session(strCookieURL & "strNTGroupsSTR"), ", ")
			for j = 0 to ubound(NTGroupSTR)
				NTGroupDBSTR = Split(fFPasswordNew, ", ")
				for i = 0 to ubound(NTGroupDBSTR)
					if NTGroupDBSTR(i) = NTGroupSTR(j) then
						chkForumAccessNew = True
						exit function
					end if
				next
			next
		case else
			chkForumAccessNew = true
	end select
end function

'##############################################
'##             RSS Functions                ##
'##############################################
function chkForumRSSAccess(fForum, UserNum, tryChkID)
	'Seems to be a modified version of chkForumAccess() originally found in inc_func_secure.asp
	'Rather than dig too far into this to see if the modifications were really that needed, I'll just rename and 
	'  use with minor changes for now
	UsrLev = chkRssUsr(UserNum, tryChkID)
	if UsrLev > 0 then
		if UsrLev < 1 then
			chkForumRSSAccess = false
		elseif UsrLev = 3 then
			chkForumRSSAccess = true
		elseif UsrLev = 4 then
			chkForumRSSAccess = true
			exit function
		end if
	else
		chkForumRSSAccess = false
	end if

	'## Forum_SQL
	strSql = "SELECT F_PRIVATEFORUMS, F_SUBJECT, F_PASSWORD_NEW "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE FORUM_ID = " & cLng(fForum)

	Set rsStatus = Server.CreateObject("ADODB.Recordset")
	rsStatus.open strSql, my_Conn

	if rsStatus.EOF or rsStatus.BOF then
		rsStatus.close
		set rsStatus = nothing
		chkForumRSSAccess = false
	else
		dim Users
		dim MatchFound
		If rsStatus("F_PRIVATEFORUMS") <> 0 then

			Select case rsStatus("F_PRIVATEFORUMS")
				case 0
					chkForumRSSAccess = true
				case 1, 3, 6 '## Allowed Users
					if OldisAllowedMember(fForum,UserNum) = 1 then
						chkForumRSSAccess = true
					else
						chkForumRSSAccess = false
					end if
				case 2 '## password
					'not supported
					chkForumRSSAccess = false
				' case 3 '## Either Password or Allowed
					'not supported unless allowed
				' case 7 '## members or password
					'not supported unless member
				case 4, 5, 7 '## members only
					if Usernum = -1 or Usernum = "" then
						chkForumRSSAccess = false
					else '## V3.1 SR4
						chkForumRSSAccess = true
					end if
				case 8, 9
					'not supported unless member
					chkForumRSSAccess = false
				case else
					chkForumRSSAccess = true
			end select
		else
			chkForumRSSAccess = true
		end if
	end if

	rsStatus.close
	set rsStatus = nothing
	
end function


sub doLoginForm()
	strMessage = "<li>You do not have access to this forum.</li>" & vbNewLine & _
				"<li>If you have been given special permission by the administrator to view and/or post in this forum, enter the password here</li>" & vbNewLine
	Call FailMessage(strMessage,False)
	
	Response.Write	"<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" id=""form2"" name=""form2"">"
	for each q in Request.QueryString
		Response.Write "<input type=""hidden"" name=""" & q & """ value=""" & Request(q) & """>" & vbNewLine
	next
	
	Response.Write	"<table class=""admin"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & vbNewLine & _
					"<input name=""pass"" type=""password"" size=""25"">" & vbNewLine & _
					"<button type=""submit"" id=""submit2"" name=""submit2"">Enter</button>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
	WriteFooter
	Response.End
end sub


sub doNotAllowed()
	Call FailMessage("<li>You do not have access to this forum.</li>",True)
	WriteFooter
	Response.End
end sub


sub doPasswordForm()
	if Request.QueryString <> "" then strRqQryString = "?" & Request.QueryString else strRqQryString = ""
	Call FailMessage("<li>You must enter the password for this forum.</li>",False)
	Response.Write	"<form action=""" & Request.ServerVariables("SCRIPT_NAME") & strRqQryString & """ method=""post"" id=""form2"" name=""form2"">" & vbNewLine
	for each q in Request.QueryString
		Response.Write "<input type=""hidden"" name=""" & q & """ value=""" & Request(q) & """>" & vbNewLine
	next
	Response.Write	"<table class=""admin"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & vbNewLine & _
					"<input name=""pass"" type=""password"" size=""25"">" & vbNewLine & _
					"<button type=""submit"" id=""submit1"" name=""submit1"">Enter</button>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
	WriteFooter
	Response.End
end sub


sub doNotLoggedInForm()
	Call FailMessage("<li>You must be logged in to enter this forum</li>",True)
	WriteFooter
	Response.End
end sub
%>
