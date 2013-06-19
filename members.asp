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
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%

if strDBNTUserName = "" then
	Response.Write	"<table class=""misc"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & vbNewLine & _
					getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member Information</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine
	
	Call FailMessage("<li>You must be logged in to view the Members List</li>",True)
	
	WriteFooter
	Response.End
end if

if trim(chkString(Request("method"),"SQLString")) <> "" then
	SortMethod = trim(chkString(Request("method"),"SQLString"))
	strSortMethod = "&method=" & SortMethod
	strSortMethod2 = "?method=" & SortMethod
end if

if trim(chkString(Request("mode"),"SQLString")) <> "" then
	strMode = trim(chkString(Request("mode"),"SQLString"))
	if strMode <> "search" then strMode = ""
end if

SearchName = trim(Request("M_NAME"))
if SearchName = "" then
	SearchName = trim(Request.Form("M_NAME"))
end if
SearchNameDisplay = Server.HTMLEncode(SearchName)
SearchName = chkString(SearchName, "sqlstring")

if Request("UserName") <> "" then
	if IsNumeric(Request("UserName")) = True then srchUName = cLng(Request("UserName")) else srchUName = "1"
end if
if Request("FirstName") <> "" then
	if IsNumeric(Request("FirstName")) = True then srchFName = cLng(Request("FirstName")) else srchFName = "0"
end if
if Request("LastName") <> "" then
	if IsNumeric(Request("LastName")) = True then srchLName = cLng(Request("LastName")) else srchLName = "0"
end if
if Request("INITIAL") <> "" then
	if IsNumeric(Request("INITIAL")) = True then srchInitial = cLng(Request("INITIAL")) else srchInitial = "0"
end if

mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)

'New Search Code
If strMode = "search"  and (srchUName = "1" or srchFName = "1" or srchLName = "1" or srchInitial = "1" ) then 
	strSql = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_LEVEL, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql = strSql & "M_AIM, M_ICQ, M_MSN, M_YAHOO, M_TITLE, M_POSTS, M_LASTPOSTDATE, M_LASTHEREDATE, M_DATE "
	strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS " 
'	if Request.querystring("link") <> "sort" then
		whereSql = " WHERE ("
		tmpSql = ""
		if srchUName = "1" then
			tmpSql = tmpSql & "M_NAME LIKE '%" & SearchName & "%' OR "
			tmpSql = tmpSql & "M_USERNAME LIKE '%" & SearchName & "%'"
		end if
		if srchFName = "1" then
			if srchUName = "1" then
					tmpSql = tmpSql & " OR "
			end if
			tmpSql = tmpSql & "M_FIRSTNAME LIKE '%" & SearchName & "%'"
		end if
		if srchLName = "1" then
			if srchFName = "1" or srchUName = "1" then 
				tmpSql = tmpSql & " OR "
			end if
			tmpSql = tmpSql & "M_LASTNAME LIKE '%" & SearchName & "%' "
		end if
		if srchInitial = "1" then 
			tmpSQL = "M_NAME LIKE '" & SearchName & "%'"
		end if

		whereSql = whereSql & tmpSql &")"
		Session(strCookieURL & "where_Sql") = whereSql
'	end if	

	if Session(strCookieURL & "where_Sql") <> "" then
		whereSql = Session(strCookieURL & "where_Sql")
	else
		whereSql = ""
	end if
	strSQL3 = whereSql
else
	'## Forum_SQL - Get all members
	strSql = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_LEVEL, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql = strSql & "M_AIM, M_ICQ, M_MSN, M_YAHOO, M_TITLE, M_POSTS, M_LASTPOSTDATE, M_LASTHEREDATE, M_DATE "
	strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS "
	if mlev = 4 then
		strSql3 = " WHERE M_NAME <> 'n/a' "
	else
		strSql3 = " WHERE M_STATUS = " & 1
	end if
end if

select case SortMethod
	case "nameasc"
		strSql4 = " ORDER BY M_NAME ASC"
	case "namedesc"
		strSql4 = " ORDER BY M_NAME DESC"
	case "levelasc"
		strSql4 = " ORDER BY M_TITLE ASC, M_NAME ASC"
	case "leveldesc"
		strSql4 = " ORDER BY M_TITLE DESC, M_NAME ASC"
	case "lastpostdateasc"
		strSql4 = " ORDER BY M_LASTPOSTDATE ASC, M_NAME ASC"
	case "lastpostdatedesc"
		strSql4 = " ORDER BY M_LASTPOSTDATE DESC, M_NAME ASC"
	case "lastheredateasc"
		if mlev = 4 or mlev = 3 then
			strSql4 = " ORDER BY M_LASTHEREDATE ASC, M_NAME ASC"
		else
			strSql4 = " ORDER BY M_POSTS DESC, M_NAME ASC"
		end if
	case "lastheredatedesc"
		if mlev = 4 or mlev = 3 then 
			strSql4 = " ORDER BY M_LASTHEREDATE DESC, M_NAME ASC"
		else
			strSql4 = " ORDER BY M_POSTS DESC, M_NAME ASC"
		end if
	case "dateasc"
		strSql4 = " ORDER BY M_DATE ASC, M_NAME ASC"
	case "datedesc"
		strSql4 = " ORDER BY M_DATE DESC, M_NAME ASC"
	case "countryasc"
		strSql4 = " ORDER BY M_COUNTRY ASC, M_NAME ASC"
	case "countrydesc"
		strSql4 = " ORDER BY M_COUNTRY DESC, M_NAME ASC"
	case "postsasc"
		strSql4 = " ORDER BY M_POSTS ASC, M_NAME ASC"
	case else
		strSql4 = " ORDER BY M_POSTS DESC, M_NAME ASC"
end select

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		OffSet = cLng((mypage - 1) * strPageSize)
		strSql5 = " LIMIT " & OffSet & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount 
	strSql1 = "SELECT COUNT(MEMBER_ID) AS PAGECOUNT "

	set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing

	if iPageTotal > 0 then
		maxpages = (iPageTotal \ strPageSize )
		if iPageTotal mod strPageSize <> 0 then
			maxpages = maxpages + 1
		end if
		if iPageTotal < (strPageSize + 1) then
			intGetRows = iPageTotal
		elseif (mypage * strPageSize) > iPageTotal then
			intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
		else
			intGetRows = strPageSize
		end if
	else
		iPageTotal = 0
		maxpages = 0
	end if 

	if iPageTotal > 0 then
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			arrMemberData = rs.GetRows(intGetRows)
			iMemberCount = UBound(arrMemberData, 2)
		rs.close
		set rs = nothing
	else
		iMemberCount = ""
	end if
else 'end MySql specific code
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
		If not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize = strPageSize
			rs.absolutepage = mypage '**
			maxpages = cLng(rs.pagecount)
			arrMemberData = rs.GetRows(strPageSize)
			iMemberCount = UBound(arrMemberData, 2)
		else
			iMemberCount = ""
		end if
	rs.Close
	set rs = nothing
end if
 
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member Information</td>" & vbNewLine & _
				"<td class=""toppaging"">" & vbNewLine
if maxpages > 1 then
	Call Paging2(1)
else
	Response.Write	"&nbsp;" & vbNewLine
end if
Response.Write	"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

Response.Write	"<table class=""admin"">" & vbNewline & _
				"<tr>" & vbNewline & _
				"<td class=""options"">" & vbNewLine & _
				"<form action=""members.asp" & strSortMethod2 & """ method=""post"" name=""SearchMembers"">" & vbNewline & _
				"<b>Search:</b>&nbsp;" & vbNewline & _
				"<input type=""checkbox"" name=""UserName"" value=""1"""
if ((srchUName <> "")  or (srchUName = "" and srchFName = "" and srchLName = "") ) then Response.Write(" checked")
Response.Write	">&nbsp;User Names" & vbNewline
if strFullName = "1" then
	Response.Write	"&nbsp;&nbsp;<input type=""checkbox"" name=""FirstName"" value=""1""" & chkCheckbox(srchFName,1,true) & ">&nbsp;First Name" & vbNewline & _
					"&nbsp;&nbsp;<input type=""checkbox"" name=""LastName"" value=""1""" & chkCheckbox(srchLName,1,true) & ">&nbsp;Last Name" & vbNewline
end if
Response.Write	"&nbsp;&nbsp;<b>For:</b>&nbsp;" & vbNewline & _
				"<input type=""text"" name=""M_NAME"" value=""" & SearchNameDisplay & """>" & vbNewline & _
				"<input type=""hidden"" name=""mode"" value=""search"">" & vbNewline & _
				"<input type=""hidden"" name=""initial"" value=""0"">" & vbNewline
if strGfxButtons = "1" then
	Response.Write	"&nbsp;&nbsp;<input src=""" & strImageUrl & "button_go.gif"" alt=""Quick Search"" type=""image"" value=""search"" id=""submit1"" name=""submit1"">" & vbNewline
else
	Response.Write	"&nbsp;&nbsp;<button type=""submit"" id=""submit1"" name=""submit1"">Search</button>" & vbNewline
end if
Response.Write	"</form>" & vbNewline & _
				"</td>" & vbNewline & _
				"</tr>" & vbNewline & _
				"<tr>" & vbNewLine & _
				"<td class=""options"">" & vbNewLine & _
				"<a href=""members.asp""" & dWStatus("Display ALL Member Names") & ">All</a>&nbsp;" & vbNewLine
for intChar = 65 to 90
	if intChar <> 90 then
		Response.Write	"<a href=""members.asp?mode=search&M_NAME=" & chr(intChar) & "&initial=1" & strSortMethod & """" & dWStatus("Display Member Names starting with the letter '" & chr(intChar) & "'") & ">" & chr(intChar) & "</a>&nbsp;" & vbNewLine
	else
		Response.Write	"<a href=""members.asp?mode=search&M_NAME=" & chr(intChar) & "&initial=1" & strSortMethod & """" & dWStatus("Display Member Names starting with the letter '" & chr(intChar) & "'") & ">" & chr(intChar) & "</a><br /></font></td>" & vbNewLine
	end if
next
Response.Write	"</tr>" & vbNewLine & _
				"</table><br />" & vbNewLine


Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine

strNames = "UserName=" & srchUName  &_
		   "&FirstName=" & srchFName &_
		   "&LastName=" & srchLName &_
		   "&INITIAL=" &srchInitial & "&"

Response.Write	"<td>&nbsp;</td>" & vbNewLine & _
				"<td><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "nameasc" then Response.Write("namedesc") else Response.Write("nameasc")
Response.Write	"""" & dWStatus("Sort by Member Name") & ">Member Name</a></td>" & vbNewLine & _
				"<td><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "levelasc" then Response.Write("leveldesc") else Response.Write("levelasc")
Response.Write	"""" & dWStatus("Sort by Member Level") & ">Title</a></td>" & vbNewLine & _
				"<td><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "postsdesc" then Response.Write("postsasc") else Response.Write("postsdesc")
Response.Write	"""" & dWStatus("Sort by Post Count") & ">Posts</a></td>" & vbNewLine & _
				"<td><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "lastpostdatedesc" then Response.Write("lastpostdateasc") else Response.Write("lastpostdatedesc")
Response.Write	"""" & dWStatus("Sort by Last Post Date") & ">Last Post</a></td>" & vbNewLine & _
				"<td><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "datedesc" then Response.Write("dateasc") else Response.Write("datedesc")
Response.Write	"""" & dWStatus("Sort by Date of Registration") & ">Member Since</a></td>" & vbNewLine
if strCountry = "1" then
	Response.Write	"<td><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
	if Request.QueryString("method") = "countryasc" then Response.Write("countrydesc") else Response.Write("countryasc")
	Response.Write	"""" & dWStatus("Sort by Country") & ">Country</a></td>" & vbNewLine
end if
if mlev = 4 or mlev = 3 then
	Response.Write	"<td><a href=""members.asp?method="
	if Request.QueryString("method") = "lastheredatedesc" then Response.Write("lastheredateasc") else Response.Write("lastheredatedesc")
	Response.Write	"""" & dWStatus("Sort by Last Visit Date") & ">Last Visit</a></td>" & vbNewLine
end if
if mlev = 4 or (lcase(strNoCookies) = "1") then
	Response.Write	"<td>&nbsp;</td>" & vbNewLine
end if
Response.Write	"</tr>" & vbNewLine

if iMemberCount = "" then '## No Members Found in DB
	Response.Write	"<tr>" & vbNewLine & _
					"<td colspan=""" & sGetColspan(9, 8) & """>No Members Found</td>" & vbNewLine & _
					"</tr>" & vbNewLine
else
	mMEMBER_ID = 0
	mM_STATUS = 1
	mM_NAME = 2
	mM_LEVEL = 3
	mM_EMAIL = 4
	mM_COUNTRY = 5
	mM_HOMEPAGE = 6
	mM_AIM = 7
	mM_ICQ = 8
	mM_MSN = 9
	mM_YAHOO = 10
	mM_TITLE = 11
	mM_POSTS = 12
	mM_LASTPOSTDATE = 13
	mM_LASTHEREDATE = 14
	mM_DATE = 15

	rec = 1
	intI = 0
	for iMember = 0 to iMemberCount
		if (rec = strPageSize + 1) then exit for

		Members_MemberID = arrMemberData(mMEMBER_ID, iMember)
		Members_MemberStatus = arrMemberData(mM_STATUS, iMember)
		Members_MemberName = arrMemberData(mM_NAME, iMember)
		Members_MemberLevel = arrMemberData(mM_LEVEL, iMember)
		Members_MemberEMail = arrMemberData(mM_EMAIL, iMember)
		Members_MemberCountry = arrMemberData(mM_COUNTRY, iMember)
		Members_MemberHomepage = arrMemberData(mM_HOMEPAGE, iMember)
		Members_MemberAIM = arrMemberData(mM_AIM, iMember)
		Members_MemberICQ = arrMemberData(mM_ICQ, iMember)
		Members_MemberMSN = arrMemberData(mM_MSN, iMember)
		Members_MemberYAHOO = arrMemberData(mM_YAHOO, iMember)
		Members_MemberTitle = arrMemberData(mM_TITLE, iMember)
		Members_MemberPosts = arrMemberData(mM_POSTS, iMember)
		Members_MemberLastPostDate = arrMemberData(mM_LASTPOSTDATE, iMember)
		Members_MemberLastHereDate = arrMemberData(mM_LASTHEREDATE, iMember)
		Members_MemberDate = arrMemberData(mM_DATE, iMember)

		if intI = 1 then 
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if

		Response.Write	"<tr>" & vbNewLine & _
						"<td class=""options"">" & vbNewLine
		if strUseExtendedProfile then
			Response.Write	"<a href=""pop_profile.asp?mode=display&id=" & Members_MemberID & """" & dWStatus("View " & ChkString(Members_MemberName,"display") & "'s Profile") & ">"
		else
			Response.Write	"<a href=""JavaScript:openWindow3('pop_profile.asp?mode=display&id=" & Members_MemberID & "')""" & dWStatus("View " & ChkString(Members_MemberName,"display") & "'s Profile") & ">"
		end if
		if Members_MemberStatus = 0 then
			Response.Write	getCurrentIcon(strIconProfileLocked,"View " & ChkString(Members_MemberName,"display") & "'s Profile","align=""absmiddle"" hspace=""0""")
		else 
			Response.Write	getCurrentIcon(strIconProfile,"View " & ChkString(Members_MemberName,"display") & "'s Profile","align=""absmiddle"" hspace=""0""")
		end if 
		Response.Write	"</a>" & vbNewLine
		if strAIM = "1" and Trim(Members_MemberAIM) <> "" then
			Response.Write	"<a href=""JavaScript:openWindow('pop_messengers.asp?mode=AIM&ID=" & Members_MemberID & "')""" & dWStatus("Send " & ChkString(Members_MemberName,"display") & " an AOL message") & ">" & getCurrentIcon(strIconAIM,"Send " & ChkString(Members_MemberName,"display") & " an AOL message","align=""absmiddle"" hspace=""0""") & "</a>" & vbNewLine
		end if
		if strICQ = "1" and Trim(Members_MemberICQ) <> "" then
			Response.Write	"<a href=""JavaScript:openWindow6('pop_messengers.asp?mode=ICQ&ID=" & Members_MemberID & "')""" & dWStatus("Send " & ChkString(Members_MemberName,"display") & " an ICQ Message") & ">" & getCurrentIcon(strIconICQ,"Send " & ChkString(Members_MemberName,"display") & " an ICQ Message","align=""absmiddle"" hspace=""0""") & "</a>" & vbNewLine
		end if
		if strMSN = "1" and Trim(Members_MemberMSN) <> "" then
			Response.Write	"<a href=""JavaScript:openWindow('pop_messengers.asp?mode=MSN&ID=" & Members_MemberID & "')""" & dWStatus("Click to see " & ChkString(Members_MemberName,"display") & "'s MSN Messenger address") & ">" & getCurrentIcon(strIconMSNM,"Click to see " & ChkString(Members_MemberName,"display") & "'s MSN Messenger address","align=""absmiddle"" hspace=""0""") & "</a>" & vbNewLine
		end if
		if strYAHOO = "1" and Trim(Members_MemberYAHOO) <> "" then
			Response.Write	"<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(Members_MemberYAHOO, "urlpath") & "&.src=pg"" target=""_blank""" & dWStatus("Send " & ChkString(Members_MemberName,"display") & " a Yahoo! Message") & ">" & getCurrentIcon(strIconYahoo,"Send " & ChkString(Members_MemberName,"display") & " a Yahoo! Message","align=""absmiddle"" hspace=""0""") & "</a>" & vbNewLine
		end if
		Response.Write	"</td>" & vbNewLine & _
						"<td>" & vbNewLine
		if strUseExtendedProfile then
			Response.Write	"<a href=""pop_profile.asp?mode=display&id=" & Members_MemberID & """ title=""View " & ChkString(Members_MemberName,"display") & "'s Profile""" & dWStatus("View " & ChkString(Members_MemberName,"display") & "'s Profile") & ">"
		else
			Response.Write	"<a href=""JavaScript:openWindow3('pop_profile.asp?mode=display&id=" & Members_MemberID & "')"" title=""View " & ChkString(Members_MemberName,"display") & "'s Profile""" & dWStatus("View " & ChkString(Members_MemberName,"display") & "'s Profile") & ">"
		end if
		Response.Write	ChkString(Members_MemberName,"display") & "</a></td>" & vbNewLine & _
						"<td>" & ChkString(getMember_Level(Members_MemberTitle, Members_MemberLevel, Members_MemberPosts),"display") & "</td>" & vbNewLine & _
						"<td class=""counts"">"
		if IsNull(Members_MemberPosts) then
			Response.Write("-")
		else
			Response.Write(Members_MemberPosts)
			if strShowRank = 2 or strShowRank = 3 then 
				Response.Write("<br />" & getStar_Level(Members_MemberLevel, Members_MemberPosts) & "")
			end if
		end if
		Response.Write	"</td>" & vbNewLine
		if IsNull(Members_MemberLastPostDate) or Trim(Members_MemberLastPostDate) = "" then
			Response.Write	"<td>-</td>" & vbNewLine
		else
			Response.Write	"<td>" & ChkDate(Members_MemberLastPostDate,"",false) & "</td>" & vbNewLine
		end if
		Response.Write	"<td>" & ChkDate(Members_MemberDate,"",false) & "</td>" & vbNewLine
		if strCountry = "1" then
			Response.Write	"<td>"
			if trim(Members_MemberCountry) <> "" then Response.Write(Members_MemberCountry & "&nbsp;") else Response.Write("-")
			Response.Write	"</td>" & vbNewLine
		end if
		if mlev = 4 or mlev = 3 then
			Response.Write	"<td>" & ChkDate(Members_MemberLastHereDate,"",false) & "</td>" & vbNewLine
		end if
		if mlev = 4 or (lcase(strNoCookies) = "1") then
			Response.Write	"<td class=""options"">" & vbNewLine
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				if Members_MemberStatus <> 0 then
					Response.Write	"<a href=""JavaScript:openWindow('pop_lock.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')""" & dWStatus("Lock Member") & ">" & getCurrentIcon(strIconLock,"Lock Member","hspace=""0""") & "</a>" & vbNewLine
					Response.Write	"<a href=""JavaScript:openWindow('pop_lock.asp?mode=Zap&MEMBER_ID=" & Members_MemberID & "')""" & dWStatus("Zap Member") & ">" & getCurrentIcon(strIconZap,"Zap Member Profile","hspace=""0""") & "</a>" & vbNewLine
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
		end if
		Response.Write	"</tr>" & vbNewLine

		rec = rec + 1
		intI = intI + 1
		if intI = 2 then intI = 0
	next
end if 
Response.Write	"</table>" & vbNewLine

if maxpages > 1 then
	Response.Write	"<table class=""misc2"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""bottompaging"">" & vbNewLine
	Call Paging2(2)
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine
end if

WriteFooter
Response.End

sub Paging2(fnum)
	if maxpages > 1 then
		if mypage = "" then
			sPageNumber = 1
		else
			sPageNumber = mypage
		end if
		If cLng(sPageNumber) > 1 then
			MinPageToShow = cLng(sPageNumber) - 1
		Else
			MinPageToShow = 1
		End If
		If cLng(sPageNumber) + strPageNumberSize > maxpages then
			MaxPageToShow = maxpages
		Else
			MaxPageToShow = cLng(sPageNumber) + strPageNumberSize
		End If
		If MaxPageToShow < maxpages then
			ShowMaxPage = True
		Else
			ShowMaxPage = False
		End If
		If MinPageToShow > 1 then
			ShowMinPage = True
		Else
			ShowMinPage = False
		End If
		
		strPageLink = "members.asp?"
		if SortMethod = "" then
			strPageLink = strPageLink & "method=postsdesc"
		else
			strPageLink = strPageLink & "method=" & SortMethod
		end if
		if srchInitial <> "" then strPageLink = strPageLink & "&initial=" & srchInitial
		if strMode <> "" then strPageLink = strPageLink & "&mode=" & strMode
		if searchName <> "" then strPageLink = strPageLink & "&M_NAME=" & searchName
		if srchUName <> "" then strPageLink = strPageLink & "&UserName=" & srchUName
		if srchFName <> "" then strPageLink = strPageLink & "&FirstName=" & srchFName
		if srchLName <> "" then strPageLink = strPageLink & "&LastName=" & srchLName		
		
		if fnum = 1 then
			Response.Write	"<b>Page:</b>&nbsp;["
		else
			Response.Write	"<b>Total Pages (" & maxpages &"):</b>&nbsp;["
		end if
		If ShowMinPage then
			Response.Write	" <a href=""" & strPageLink & "&whichpage=1""><< First</a> "
		End If
		For counter = MinPageToShow To MaxPageToShow
			if counter <> cLng(pge) then   
				Response.Write	" <a href=""" & strPageLink & "&whichpage=" & counter & """>" & counter & "</a> "
			else
				Response.Write	" <b>[" & counter & "]</b>"
			end if
		Next
		If ShowMaxPage then
			Response.Write	" <a href=""" & strPageLink & "&whichpage=" & maxpages & """>>> Last</a> "
		End If
		Response.Write	"]"
	end if
end sub 

Function sGetColspan(lIN, lOUT)
	if (mlev = "4" or mlev = "3") then lOut = lOut + 2
	If lOut > lIn then
		sGetColspan = lIN
	Else
		sGetColspan = lOUT
	End If
end Function
%>
