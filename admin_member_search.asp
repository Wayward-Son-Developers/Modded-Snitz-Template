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
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if

SearchMember = trim(chkString(Request.Form("SearchMember"),"SQLString"))
SearchType = cLng(Request.Form("SearchType"))

select case SearchType
	case 1 : strSql3 = " WHERE M_NAME LIKE '%" & SearchMember & "%'"
	case 2 : strSql3 = " WHERE M_EMAIL LIKE '%" & SearchMember & "%'"
	case 3 : strSql3 = " WHERE M_IP LIKE '%" & SearchMember & "%' OR M_LAST_IP LIKE '%" & SearchMember & "%'"
	case else : strSql3 = " WHERE M_NAME LIKE '%" & SearchMember & "%'"
end select

if Request.QueryString("mode") = "DoIt" then
	
	Response.Write	"<table class=""misc"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""secondnav"">" & vbNewLine & _
					getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;<a href=""admin_member_search.asp"">Member&nbsp;Search</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Search&nbsp;Results</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine

	if SearchMember <> "" then

		'## Forum_SQL - Get all members
		strSql = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_LEVEL, M_EMAIL, M_TITLE, M_POSTS, M_LASTPOSTDATE, M_LASTHEREDATE, M_DATE, M_IP, M_LAST_IP "
		strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql4 = " ORDER BY M_LEVEL DESC, M_NAME ASC"

		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			if rs.EOF then
				iMemberCount = ""
			else
				arrMemberData = rs.GetRows(adGetRowsRest)
				iMemberCount = UBound(arrMemberData,2)
			end if
		rs.Close
		set rs = nothing

		Response.Write	"<table class=""admin"" width=""100%"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td>UserName</td>" & vbNewLine & _
						"<td>Title</td>" & vbNewLine & _
						"<td>E-Mail Address</td>" & vbNewLine & _
						"<td>IP Address</td>" & vbNewLine & _
						"<td>Last IP Address</td>" & vbNewLine & _
						"<td>Member Since</td>" & vbNewLine & _
						"<td>Last Post</td>" & vbNewLine & _
						"<td>Last Visit</td>" & vbNewLine & _
						"<td>&nbsp;</td>" & vbNewLine & _
						"</tr>" & vbNewLine
		if iMemberCount = "" then '## No Members Found in DB
			Response.Write	"<tr>" & vbNewLine & _
							"<td colspan=""9"">No Members Found</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		else
			mMEMBER_ID = 0
			mM_STATUS = 1
			mM_NAME = 2
			mM_LEVEL = 3
			mM_EMAIL = 4
			mM_TITLE = 5
			mM_POSTS = 6
			mM_LASTPOSTDATE = 7
			mM_LASTHEREDATE = 8
			mM_DATE = 9
			mM_IP = 10
			mM_LAST_IP = 11

			intI = 0

			for iMember = 0 to iMemberCount
				Members_MemberID = arrMemberData(mMEMBER_ID, iMember)
				Members_MemberStatus = arrMemberData(mM_STATUS, iMember)
				Members_MemberName = arrMemberData(mM_NAME, iMember)
				Members_MemberLevel = arrMemberData(mM_LEVEL, iMember)
				Members_MemberEMail = arrMemberData(mM_EMAIL, iMember)
				Members_MemberTitle = arrMemberData(mM_TITLE, iMember)
				Members_MemberPosts = arrMemberData(mM_POSTS, iMember)
				Members_MemberLastPostDate = arrMemberData(mM_LASTPOSTDATE, iMember)
				Members_MemberLastHereDate = arrMemberData(mM_LASTHEREDATE, iMember)
				Members_MemberDate = arrMemberData(mM_DATE, iMember)
				Members_MemberIP = arrMemberData(mM_IP, iMember)
				Members_MemberLastIP = arrMemberData(mM_LAST_IP, iMember)

				if intI = 1 then 
					CColor = strAltForumCellColor
				else
					CColor = strForumCellColor
				end if

				Response.Write	"<tr>" & vbNewLine & _
								"<td" & chkHiLite(SearchType,1) & ">" & profileLink(ChkString(Members_MemberName,"display"),Members_MemberID) & "</td>" & vbNewLine & _
								"<td>" & ChkString(getMember_Level(Members_MemberTitle, Members_MemberLevel, Members_MemberPosts),"display") & "</td>" & vbNewLine & _
								"<td" & chkHiLite(SearchType,2) & ">" & ChkString(Members_MemberEmail,"display") & "</td>" & vbNewLine & _
								"<td" & chkHiLite(SearchType,3) & "><a href=""" & strIPLookup & ChkString(Members_MemberIP,"display") & """ target=""_blank"">" & ChkString(Members_MemberIP,"display") & "</a></td>" & vbNewLine & _
								"<td" & chkHiLite(SearchType,3) & "><a href=""" & strIPLookup & ChkString(Members_MemberLastIP,"display") & """ target=""_blank"">" & ChkString(Members_MemberLastIP,"display") & "</a></td>" & vbNewLine
				Response.Write	"<td class=""counts"">" & ChkDate(Members_MemberDate,"",false) & "</td>" & vbNewLine
				if IsNull(Members_MemberLastHereDate) or Trim(Members_MemberLastPostDate) = "" then
					Response.Write	"<td class=""counts"">-</td>" & vbNewLine
				else
					Response.Write	"<td class=""counts""><span class=""lastpost"">" & ChkDate(Members_MemberLastPostDate,"<br />",true) & "</span></td>" & vbNewLine
				end if
				Response.Write	"<td class=""counts""><span class=""lastpost"">" & ChkDate(Members_MemberLastHereDate,"<br />",true) & "</span></td>" & vbNewLine
				Response.Write	"<td class=""options"">" & vbNewLine
				if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
					'## Do Nothing
				else
					if Members_MemberStatus <> 0 then
						Response.Write	"<a href=""JavaScript:openWindow('pop_lock.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')""" & dWStatus("Lock Member") & ">" & getCurrentIcon(strIconLock,"Lock Member","hspace=""0""") & "</a>" & vbNewLine
					else
						Response.Write	"<a href=""JavaScript:openWindow('pop_open.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')""" & dWStatus("Un-Lock Member") & ">" & getCurrentIcon(strIconUnlock,"Un-Lock Member","hspace=""0""") & "</a>" & vbNewLine
					end if
				end if
				if (Members_MemberID = intAdminMemberID and MemberID <> intAdminMemberID) OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID AND MemberID <> Members_MemberID) then
					Response.Write	"-" & vbNewLine
				else
					if strUseExtendedProfile then
						Response.Write	"<a href=""pop_profile.asp?mode=Modify&ID=" & Members_MemberID & """" & dWStatus("Edit Member") & ">" & getCurrentIcon(strIconPencil,"Edit Member","hspace=""0""") & "</a>" & vbNewLine
					else
						Response.Write	"<a href=""JavaScript:openWindow3('pop_profile.asp?mode=Modify&ID=" & Members_MemberID & "')""" & dWStatus("Edit Member") & ">" & getCurrentIcon(strIconPencil,"Edit Member","hspace=""0""") & "</a>" & vbNewLine
					end if
				end if
				if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
					'## Do Nothing
				else
					Response.Write	"<a href=""JavaScript:openWindow('pop_delete.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')""" & dWStatus("Delete Member") & ">" & getCurrentIcon(strIconTrashcan,"Delete Member","hspace=""0""") & "</a>" & vbNewLine
				end if
				Response.Write	"</td>" & vbNewLine
				Response.Write	"</tr>" & vbNewLine

				intI = intI + 1
				if intI = 2 then intI = 0
			next
		end if 
		Response.Write	"</table>" & vbNewLine
	else
		Call FailMessage("<li>You must enter something to search for</li>",True)
	end if
else
	Response.Write	"<table class=""misc"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""secondnav"">" & vbNewLine & _
					getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member&nbsp;Search</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine

	Response.Write	"<form action=""admin_member_search.asp?mode=DoIt"" name=""SearchForm"" id=""SearchForm"" method=""post"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Member Search</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">" & _
					"<select name=""searchType"">" & vbNewLine & _
					"<option value=""1"" selected>User Name contains...</option>" & vbNewLine & _
					"<option value=""2"">E-Mail Address contains...</option>" & vbNewLine & _
					"<option value=""3"">IP Address contains...</option>" & vbNewLine & _
					"</select></td>" & vbNewLine & _
					"<td class=""formvalue""><input name=""SearchMember"" value="""" size=""50""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"">Find Member</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if 
WriteFooter
Response.End

function chkHiLite(actualValue, thisValue)
	if isNumeric(actualValue) then actualValue = cLng(actualValue)
	chkHiLite = " class=""counts"
	if actualValue = thisValue then
		chkHiLite = chkHiLite & " foundword"
	end if
	chkHiLite = chkHiLite & """"
end function
%>
