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
<%
Dim HasHigherSub
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts
Dim canView
HasHigherSub = false
Dim strUseMemberDropDownBox
strUseMemberDropDownBox = 0
%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_chknew.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<% 
Dim ArchiveView
Dim AdminAllowed, ModerateAllowed
Dim SearchLink : SearchLink = ""
if request("ARCHIVE") = "true" then
	strActivePrefix = strArchiveTablePrefix
	ArchiveView = "true"
	ArchiveLink = "ARCHIVE=true&"
else
	strActivePrefix = strTablePrefix
	ArchiveView = ""
	ArchiveLink = ""
end if

select case cLng(Request.Form("andor"))
	case 1 : strAndOr = " and "
	case 2 : strAndOr = " or "
	case 3 : strAndOr = "phrase"
	case else : strAndOr = " and "
end select

Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
				"    function ChangePage(fnum){" & vbNewLine & _
				"    	if (fnum == 1) {" & vbNewLine & _
				"    		document.PageNum1.submit();" & vbNewLine & _
				"    		}" & vbNewLine & _
				"    	else {" & vbNewLine & _
				"    		document.PageNum2.submit();" & vbNewLine & _
				"    	}" & vbNewLine & _
				"    }" & vbNewLine & _
				"</script>" & vbNewLine

' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs, strTopicSubs
if MySubCount > 0 then
	strSubString = PullSubscriptions(0,0,0)
	strSubArray  = Split(strSubString,";")
	if uBound(strSubArray) < 0 then
		strBoardSubs = ""
		strCatSubs = ""
		strForumSubs = ""
		strTopicSubs = ""
	else
		strBoardSubs = strSubArray(0)
		strCatSubs = strSubArray(1)
		strForumSubs = strSubArray(2)
		strTopicSubs = strSubArray(3)
	end if
end if

' DEM --> Added code for topic moderation
if mlev = 3 then
	strSql = "SELECT FORUM_ID FROM " & strTablePrefix & "MODERATOR " & _
		 " WHERE MEMBER_ID = " & MemberID

	Set rsMod = Server.CreateObject("ADODB.Recordset")
	rsMod.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rsMod.EOF then
		recModCount = ""
	else
		allModData = rsMod.GetRows(adGetRowsRest)
		recModCount = UBound(allModData,2)
	end if

	RsMod.close
	set RsMod = nothing

	if recModCount <> "" then
		for x = 0 to recModCount
			if x = 0 then
				ModOfForums = allModData(0,x)
			else
				ModOfForums = ModOfForums & "," & allModData(0,x)
			end if
		next
	else
		ModOfForums = ""
	end if
else
	ModOfForums = ""
end if

if strPrivateForums = "1" and mLev < 4 then
	allAllowedForums = ""

	allowSql = "SELECT FORUM_ID, F_SUBJECT, F_PRIVATEFORUMS, F_PASSWORD_NEW"
	allowSql = allowSql & " FROM " & strTablePrefix & "FORUM"
	allowSql = allowSql & " WHERE F_TYPE = 0"
	allowSql = allowSql & " ORDER BY FORUM_ID"

	set rsAllowed = Server.CreateObject("ADODB.Recordset")
	rsAllowed.open allowSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rsAllowed.EOF then
		recAllowedCount = ""
	else
		allAllowedData = rsAllowed.GetRows(adGetRowsRest)
		recAllowedCount = UBound(allAllowedData,2)
	end if

	rsAllowed.close
	set rsAllowed = nothing

	if recAllowedCount <> "" then
		fFORUM_ID = 0
		fF_SUBJECT = 1
		fF_PRIVATEFORUMS = 2
		fF_PASSWORD_NEW = 3

		for RowCount = 0 to recAllowedCount

			Forum_ID = allAllowedData(fFORUM_ID,RowCount)
			Forum_Subject = allAllowedData(fF_SUBJECT,RowCount)
			Forum_PrivateForums = allAllowedData(fF_PRIVATEFORUMS,RowCount)
			Forum_FPasswordNew = allAllowedData(fF_PASSWORD_NEW,RowCount)

			if mLev = 4 then
				ModerateAllowed = "Y"
			elseif mLev = 3 and ModOfForums <> "" then
				if (strAuthType = "nt") then
					if (chkForumModerator(Forum_ID, Session(strCookieURL & "username")) = "1") then ModerateAllowed = "Y" else ModerateAllowed = "N"
				else 
					if (instr("," & ModOfForums & "," ,"," & Forum_ID & ",") > 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
				end if
			else
				ModerateAllowed = "N"
			end if
			if chkForumAccessNew(Forum_PrivateForums,Forum_FPasswordNew,Forum_Subject,Forum_ID,MemberID) = true then
				if allAllowedForums = "" then
					allAllowedForums = Forum_ID
				else
					allAllowedForums = allAllowedForums & "," & Forum_ID
				end if
			end if
		next
	end if
	if allAllowedForums = "" then allAllowedForums = 0
end if

if Request.QueryString("mode") = "DoIt" then
	if trim(Request.Form("Search")) <> "" or Request.QueryString("MEMBER_ID") <> "" then
		if Request.Form("Search") <> "" then
			keywords = split(Request.Form("Search"), " ")
			keycnt = ubound(keywords)
			for i = 0 to keycnt
				if i = 0 then
					strKeywords = keywords(i)
				else
					strKeywords = strKeywords & "," & keywords(i)
				end if
			next
			if strAndOr = "phrase" then strKeyWords = replace(strKeyWords,",","+")
			SearchLink = "&SearchTerms=" & chkString(strKeyWords,"search")
		end if

		'## Forum_SQL - Find all records with the search criteria in them
		strSql = "SELECT DISTINCT C.CAT_STATUS, C.CAT_SUBSCRIPTION, C.CAT_NAME, C.CAT_ORDER"
		strSql = strSql & ", F.F_ORDER, F.FORUM_ID, F.F_SUBJECT, F.CAT_ID"
		strSql = strSql & ", F.F_SUBSCRIPTION, F.F_STATUS"
		strSql = strSql & ", T.TOPIC_ID, T.T_AUTHOR, T.T_SUBJECT, T.T_STATUS, T.T_LAST_POST"
		strSql = strSql & ", T.T_LAST_POST_AUTHOR, T.T_LAST_POST_REPLY_ID, T.T_REPLIES, T.T_UREPLIES, T.T_VIEW_COUNT"
		strSql = strSql & ", M.MEMBER_ID, M.M_NAME, MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME"
		strSql = strSql & ", F.F_PRIVATEFORUMS, F.F_PASSWORD_NEW"

		strSql2 = " FROM ((((" & strTablePrefix & "FORUM F LEFT JOIN " & strActivePrefix & "TOPICS T"
		strSql2 = strSql2 & " ON F.FORUM_ID = T.FORUM_ID) LEFT JOIN " & strActivePrefix & "REPLY R"
		strSql2 = strSql2 & " ON T.TOPIC_ID = R.TOPIC_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS M"
		strSql2 = strSql2 & " ON T.T_AUTHOR = M.MEMBER_ID) LEFT JOIN " & strTablePrefix & "CATEGORY C"
		strSql2 = strSql2 & " ON T.CAT_ID = C.CAT_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS MEMBERS_1"
		strSql2 = strSql2 & " ON T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID"
		if Request.Form("Search") <> "" then
			strSql3 = " WHERE ("
			'################# New Search Code #################################################
			if Request.Form("SearchMessage") = 1 then
				if strAndOr = "phrase" then
					strSql3 = strSql3 & " (T.T_SUBJECT LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%') "
				else
					For Each word in keywords
						SearchWord = ChkString(word, "SQLString")
						strSql3 = strSql3 & " (T.T_SUBJECT LIKE '%" & SearchWord & "%') "
						if cnt < keycnt then strSql3 = strSql3 & strAndOr
						cnt = cnt + 1
					next
				end if
			else
				if strAndOr = "phrase" then
					strSql3 = strSql3 & " (R.R_MESSAGE LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%'"
					strSql3 = strSql3 & " OR T.T_SUBJECT LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%'"
					strSql3 = strSql3 & " OR T.T_MESSAGE LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%') "
				else
					For Each word in keywords
						SearchWord = ChkString(word, "SQLString")
						strSql3 = strSql3 & " (R.R_MESSAGE LIKE '%" & SearchWord & "%'"
						strSql3 = strSql3 & " OR T.T_SUBJECT LIKE '%" & SearchWord & "%'"
						strSql3 = strSql3 & " OR T.T_MESSAGE LIKE '%" & SearchWord & "%') "
						if cnt < keycnt then strSql3 = strSql3 & strAndOr
						cnt = cnt + 1
					next
				end if
			'################# New Search Code #################################################
			end if
			strSql3 = strSql3 & " ) "
		else
			strSql3 = " WHERE (0 = 0)"
		end if
		' DEM --> Added code to ignore unmoderated/held posts...
		if mlev <> 4 then 
			strSql3 = strSql3 & " AND ((T.T_AUTHOR <> " & MemberID
			strSql3 = strSql3 & " AND T.T_STATUS < 2)"  ' Ignore unapproved/held posts
			strSql3 = strSql3 & " OR T.T_AUTHOR = " & MemberID & ")"
		end if
		' DEM --> End of Code added to ignore unmoderated/held posts....
		cnt = 0
		if cLng(Request.Form("Forum")) <> 0 then
			strSql3 = strSql3 & " AND F.FORUM_ID = " & cLng(Request.Form("Forum")) & " "
		end if
		if cLng(Request.Form("SearchDate")) <> 0 then
			dt = cLng(Request.Form("SearchDate"))
			strSql3 = strSql3 & " AND ( (T.T_DATE > '" & DateToStr(dateadd("d", -dt, strForumTimeAdjust)) & "') " 
			strSql3 = strSql3 & "OR (R.R_DATE > '" & DateToStr(dateadd("d", -dt, strForumTimeAdjust)) & "') ) "
		end if
		if strUseMemberDropDownBox = 0 then
			intSearchMember = getMemberID(chkString(Request.Form("SearchMember"),"SQLString"))
			if intSearchMember <> 0 then
				strSql3 = strSql3 & " AND (M.MEMBER_ID = " & cLng(intSearchMember) & " "
				strSql3 = strSql3 & " OR R.R_AUTHOR = " & cLng(intSearchMember) & ") "
			end if
		else
			if cLng(Request.Form("SearchMember")) <> 0 then
				strSql3 = strSql3 & " AND (M.MEMBER_ID = " & cLng(Request.Form("SearchMember")) & " "
				strSql3 = strSql3 & " OR R.R_AUTHOR = " & cLng(Request.Form("SearchMember")) & ") "
			end if
		end if
		if cLng(Request.QueryString("MEMBER_ID")) <> 0 then
			strSql3 = strSql3 & " AND (M.MEMBER_ID = " & cLng(Request.QueryString("MEMBER_ID")) & " "
			strSql3 = strSql3 & " OR R.R_AUTHOR = " & cLng(Request.QueryString("MEMBER_ID")) & ") "
		end if
		if strPrivateForums = "1" and mLev < 4 then
			strSql3 = strSql3 & " AND F.FORUM_ID IN (" & allAllowedForums & ")"
		end if
		strSql3 = strSql3 & " AND F.F_TYPE = 0"

		strSql4 = " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT, T.T_LAST_POST DESC"

		mypage = request("whichpage")
		if ((Trim(mypage) = "") or (IsNumeric(mypage) = False)) then mypage = 1
		mypage = cLng(mypage)

		if strDBType = "mysql" then 'MySql specific code
			if mypage > 1 then 
				intOffset = cLng((mypage-1) * strPageSize)
				strSql5 = strSql5 & " LIMIT " & intOffset & ", " & strPageSize & " "
			end if

			'## Forum_SQL - Get the total pagecount 
			strSql1 = "SELECT COUNT(DISTINCT T.TOPIC_ID) AS PAGECOUNT "

			set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
			iPageTotal = rsCount(0).value
			rsCount.close
			set rsCount = nothing

			if iPageTotal > 0 then
				inttotaltopics = iPageTotal
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
				inttotaltopics = iPageTotal
				maxpages = 0
			end if 

			if iPageTotal > 0 then
				set rs = Server.CreateObject("ADODB.Recordset")
				rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
					arrTopicData = rs.GetRows(intGetRows)
					iTopicCount = UBound(arrTopicData, 2)
				rs.close
				set rs = nothing
			else
				iTopicCount = ""
			end if
 
		else 'end MySql specific code

			set rs = Server.CreateObject("ADODB.Recordset")
			rs.cachesize = strPageSize

			rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
				if not rs.EOF then
					rs.movefirst
					rs.pagesize = strPageSize
					inttotaltopics = cLng(rs.recordcount)
					rs.absolutepage = mypage '**
					maxpages = cLng(rs.pagecount)
					arrTopicData = rs.GetRows(strPageSize)
					iTopicCount = UBound(arrTopicData, 2)
				else
					iTopicCount = ""
					inttotaltopics = 0
				end if
			rs.Close
			set rs = nothing
		end if

		if strModeration = "1" and mLev > 2 then
			UnModeratedPosts = CheckForUnmoderatedPosts("BOARD", 0, 0, 0)
			UnModeratedFPosts = 0
		end if

		Response.Write	"<table class=""misc"">" & vbNewline & _
						"<tr>" & vbNewline & _
						"<td class=""secondnav"">" & vbNewLine & _
						getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewline & _
						getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""search.asp"">Search Form</a>" & vbNewLine
		if Request.Form("Search") <> "" then
			Response.Write	"<br />" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Search Results for: " & chkString(Request.Form("Search"),"display")
		elseif Request.QueryString("MEMBER_ID") <> "" then
			Response.Write	"<br />" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Search Results for: All non-archived Topics that contain posts by " & getMemberName(cLng(Request.QueryString("MEMBER_ID")))
		end if
		Response.Write	"</td>" & vbNewline & _
						"</tr>" & vbNewline
		if maxpages > 1 then
			Response.Write	"<tr align=""right"">" & vbNewLine
			Call DropDownPaging(1)
			Response.Write	"</tr>" & vbNewLine
		end if
		Response.Write	"</table>" & vbNewLine

		Response.Write	"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
						"<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&nbsp;</font></b></td>" & vbNewLine & _
						"<td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Topic</font></b></td>" & vbNewLine & _
						"<td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Author</font></b></td>" & vbNewLine & _
						"<td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Replies</font></b></td>" & vbNewLine & _
						"<td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Read</font></b></td>" & vbNewLine & _
						"<td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Last Post</font></b></td>" & vbNewLine & _
						"</tr>" & vbNewLine
		if iTopicCount = "" then '## No Search Results
			Response.Write	"<tr>" & vbNewLine & _
							"<td bgcolor=""" & strForumCellColor & """ colspan=""6""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>No Matches Found</b></font></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		else 
			cCAT_STATUS = 0
			cCAT_SUBSCRIPTION = 1
			cCAT_NAME = 2
			fFORUM_ID = 5
			fF_SUBJECT = 6
			fCAT_ID = 7
			fF_SUBSCRIPTION = 8
			fF_STATUS = 9
			tTOPIC_ID = 10
			tT_AUTHOR = 11
			tT_SUBJECT = 12
			tT_STATUS = 13
			tT_LAST_POST = 14
			tT_LAST_POST_AUTHOR = 15
			tT_LAST_POST_REPLY_ID = 16
			tT_REPLIES = 17
			tT_UREPLIES = 18
			tT_VIEW_COUNT = 19
			mMEMBER_ID = 20
			mM_NAME = 21
			tLAST_POST_AUTHOR_NAME = 22
			fF_PRIVATEFORUMS = 23
			fF_PASSWORD_NEW = 24

			currForum = 0 
			currTopic = 0
			dim Cat_Status
			dim Cat_Subscription
			dim Forum_Status
			dim Forum_Subscription
			dim mdisplayed
			mdisplayed = 0
			rec = 1

			for iTopic = 0 to iTopicCount
				if (rec = strPageSize + 1) then exit for

				Cat_Status = arrTopicData(cCAT_STATUS, iTopic)
				Cat_Subscription = arrTopicData(cCAT_SUBSCRIPTION, iTopic)
				Cat_Name = arrTopicData(cCAT_NAME, iTopic)
				Forum_ID = arrTopicData(fFORUM_ID, iTopic)
				Forum_Subject = arrTopicData(fF_SUBJECT, iTopic)
				Forum_Cat_ID = arrTopicData(fCAT_ID, iTopic)
				Forum_Subscription = arrTopicData(fF_SUBSCRIPTION, iTopic)
				Forum_Status = arrTopicData(fF_STATUS, iTopic)
				Topic_ID = arrTopicData(tTOPIC_ID, iTopic)
				Topic_Author = arrTopicData(tT_AUTHOR, iTopic)
				Topic_Subject = arrTopicData(tT_SUBJECT, iTopic)
				Topic_Status = arrTopicData(tT_STATUS, iTopic)
				Topic_LastPost = arrTopicData(tT_LAST_POST, iTopic)
				Topic_LastPostAuthor = arrTopicData(tT_LAST_POST_AUTHOR, iTopic)
				Topic_LastPostReplyID = arrTopicData(tT_LAST_POST_REPLY_ID, iTopic)
				Topic_Replies = arrTopicData(tT_REPLIES, iTopic)
				Topic_UReplies = arrTopicData(tT_UREPLIES, iTopic)
				Topic_ViewCount = arrTopicData(tT_VIEW_COUNT, iTopic)
				Topic_MemberID = arrTopicData(mMEMBER_ID, iTopic)
				Topic_MemberName = arrTopicData(mM_NAME, iTopic)
				Topic_LastPostAuthorName = arrTopicData(tLAST_POST_AUTHOR_NAME, iTopic)
				Forum_PrivateForums = arrTopicData(fF_PRIVATEFORUMS, iTopic)
				Forum_FPasswordNew = arrTopicData(fF_PASSWORD_NEW, iTopic)

				if mLev = 4 then
					AdminAllowed = 1
				else
				    	AdminAllowed = 0
				end if
				if mLev = 4 then
					ModerateAllowed = "Y"
				elseif mLev = 3 and ModOfForums <> "" then
					if (strAuthType = "nt") then
						if (chkForumModerator(Forum_ID, Session(strCookieURL & "username")) = "1") then ModerateAllowed = "Y" else ModerateAllowed = "N"
					else 
						if (instr("," & ModOfForums & "," ,"," & Forum_ID & ",") > 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
					end if
				else
					ModerateAllowed = "N"
				end if

				if ModerateAllowed = "Y" and Topic_UReplies > 0 then
					Topic_Replies = Topic_Replies + Topic_UReplies
				end if
				if (currForum <> Forum_ID) and (currTopic <> Topic_ID) then 
					Response.Write	"<tr>" & vbNewLine & _
									"<td height=""20"" colspan=""6"" bgcolor=""" & strCategoryCellColor & """ valign=""middle""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><a href=""default.asp?CAT_ID=" & Forum_Cat_ID & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>" & ChkString(Cat_Name,"display") & "</b></font></a>&nbsp;/&nbsp;<a href=""forum.asp?FORUM_ID=" & Forum_ID & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>" & ChkString(Forum_Subject,"display") & "</b></font></a></font></td>" & vbNewline & _
									"</tr>" & vbNewLine
					currForum = Forum_ID
				end if 
				if currTopic <> Topic_ID then
					Response.Write	"<tr>" & vbNewline
					if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then 
						' DEM --> Added if statement to display topic status properly
						if Topic_Status = 2 then
							UnApprovedFound = "Y"
							Response.Write	"<td bgcolor=""" & strForumCellColor & """ align=""center""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & SearchLink & """>" & getCurrentIcon(strIconFolderUnmoderated,"Topic UnModerated","hspace=""0""") & "</a></td>" & vbNewline
						elseif Topic_Status = 3 then
							HeldFound = "Y"
							Response.Write	"<td bgcolor=""" & strForumCellColor & """ align=""center""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & SearchLink & """>" & getCurrentIcon(strIconFolderHold,"Topic Held","hspace=""0""") & "</a></td>" & vbNewline
						else
							Response.Write	"<td bgcolor=""" & strForumCellColor & """ align=""center""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & SearchLink & """>" & ChkIsNew(Topic_LastPost) & "</a></td>" & vbNewline
						end if
					else 
						if Cat_Status = 0 then 
							strAltText = "Category Locked"
						elseif Forum_Status = 0 then 
							strAltText = "Forum Locked"
						else
							strAltText = "Topic Locked"
						end if 
						Response.Write	"<td bgcolor=""" & strForumCellColor & """ align=""center""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & SearchLink & """>" & getCurrentIcon(strIconFolderLocked,strAltText,"hspace=""0""") & "</a></td>" & vbNewline
					end if
					Response.Write	"<td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					Response.Write	"<span class=""spnMessageText""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & SearchLink & """>" & ChkString(left(Topic_Subject, 50),"display") & "</a></span>&nbsp;</font>" & vbNewLine
					if strShowPaging = "1" then
						TopicPaging()
					end if
					Response.Write	"</td>" & vbNewLine & _
									"<td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"">" & profileLink(chkString(Topic_MemberName,"display"),Topic_MemberID) & "</span></font></td>" & vbNewLine & _
									"<td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & Topic_Replies & "</font></td>" & vbNewLine & _
									"<td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & Topic_ViewCount & "</font></td>" & vbNewLine
					if IsNull(Topic_LastPostAuthor) then
						strLastAuthor = ""
					else
						strLastAuthor = "<br />by: <span class=""spnMessageText"">" & profileLink(Topic_LastPostAuthorName,Topic_LastPostAuthor) & "</span>"
						if (strJumpLastPost = "1" and ArchiveView = "") then strLastAuthor = strLastAuthor & "&nbsp;" & DoLastPostLink
					end if
					Response.Write	"<td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center"" nowrap><font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strFooterFontSize & """><b>" & ChkDate(Topic_LastPost, "</b>&nbsp" ,true) & strLastAuthor & "</font></td>" & vbNewLine
					Response.Write	"</tr>" & vbNewLine
					currTopic = Topic_ID
					rec = rec + 1 
				end if 
				mdisplayed = mdisplayed + 1
			next
			if mdisplayed = 0 then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgcolor=""" & strForumCellColor & """ colspan=""6""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>No Matches Found</b></font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
		end if 
		Response.Write	"</table>" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine
		Response.Write	"</table>" & vbNewLine

		if maxpages > 1 then
			Response.Write	"<table border=""0"" width=""100%"" align=""center"">" & vbNewline & _
							"<tr>" & vbNewLine
			Call DropDownPaging(2)
			Response.Write	"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		end if

		Response.Write	"<table width=""100%"" align=""center"" border=""0"">" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td align=""left"" valign=""top"">" & vbNewLine & _
						"<table>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td nowrap>" & vbNewLine & _
						"<p><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
						"" & getCurrentIcon(strIconFolderNew,"New Posts","align=""absmiddle""") & " New posts since last logon.<br />" & vbNewLine & _
						"" & getCurrentIcon(strIconFolder,"Old Posts","align=""absmiddle""") & " Old Posts."
		if lcase(strHotTopic) = "1" then Response.Write	(" (" & getCurrentIcon(strIconFolderHot,"Hot Topic","align=""absmiddle""") & "&nbsp;" & intHotTopicNum & " replies or more.)<br />" & vbNewLine)
		Response.Write	"" & getCurrentIcon(strIconFolderLocked,"Locked Topic","align=""absmiddle""") & " Locked topic.<br />" & vbNewLine
		' DEM --> Start of Code added for moderation
		if HeldFound = "Y" then
			Response.Write "" & getCurrentIcon(strIconFolderHold,"Held Topic","align=""absmiddle""") & " Held Topic.<br />" & vbNewline
		end if
		if UnApprovedFound = "Y" then
			Response.Write "" & getCurrentIcon(strIconFolderUnmoderated,"UnModerated Topic","align=""absmiddle""") & " UnModerated Topic.<br />" & vbNewline
		end if
		' DEM --> End of Code added for moderation
		Response.Write	"</font></p></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine

	else
		Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You must enter keywords</font></p>" & vbNewLine & _
						"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Back To Search Page</a></font></p>" & vbNewLine & _
						"<meta http-equiv=""Refresh"" content=""2; URL=JavaScript:history.go(-1)"">" & vbNewLine
	end if
else
	strRqForumID = cLng(Request.QueryString("FORUM_ID"))

	Response.Write	"<table class=""misc"">" & vbNewline & _
					"<tr>" & vbNewline & _
					"<td class=""secondnav"">" & vbNewLine & _
					getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewline & _
					getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;Search Form</td>" & vbNewline & _
					"</tr>" & vbNewline & _
					"</table>" & vbNewLine

	Response.Write	"<form action=""search.asp?mode=DoIt"" name=""SearchForm"" id=""SearchForm"" method=""post"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Search For:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""Search"" size=""40"" value=""" & trim(ChkString(Request.QueryString("Search"),"display")) & """><br />" & vbNewLine & _
					"<input type=""radio"" class=""radio"" name=""andor"" value=""3"">Match exact phrase<br />" & vbNewLine & _
					"<input type=""radio"" class=""radio"" name=""andor"" value=""1"" checked>Search for all Words<br />" & vbNewLine & _
					"<input type=""radio"" class=""radio"" name=""andor"" value=""2"">Match any of the words</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Search Forum:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<select name=""Forum"" size=""1"">" & vbNewLine & _
					"<option value=""0"">All Forums</option>" & vbNewLine
	'## Forum_SQL
	strSql = "SELECT F.FORUM_ID, F.F_SUBJECT FROM " & strTablePrefix & "FORUM F, " & strTablePrefix & "CATEGORY C"
	strSql = strSql & " WHERE F_TYPE = " & 0
	if strPrivateForums = "1" and allAllowedForums <> "" and mLev < 4 then
		strSql = strSql & " AND F.FORUM_ID IN (" & allAllowedForums & ")"
	end if
	strSql = strSql & " AND C.CAT_ID = F.CAT_ID"
	strSql = strSql & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT"

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recForumCount = ""
	else
		allForumData = rs.GetRows(adGetRowsRest)
		recForumCount = UBound(allForumData,2)
		fFORUM_ID = 0
		fF_SUBJECT = 1
	end if

	rs.close
	set rs = nothing

	if recForumCount <> "" then
		for iForum = 0 to recForumCount
			ForumForumID = allForumData(fFORUM_ID, iForum)
			ForumSubject = allForumData(fF_SUBJECT, iForum)
			Response.Write	"<option value=""" & ForumForumID & """"
			if strRqForumID = ForumForumID then Response.Write(" selected")
			Response.Write	">" & ChkString(left(ForumSubject, 50),"display") & "</option>" & vbNewline
		next
	end if
	Response.Write	"</select>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	'################# New Search Code #################################################
	Response.Write	"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Search In:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<select name=""SearchMessage"">" & vbNewLine & _
					"<option value=""0"">Entire Message</option>" & vbNewLine & _
					"<option value=""1"">Subject Only</option>" & vbNewLine & _
					"</select>" & vbNewLine
	if strArchiveState = "1" then Response.Write("&nbsp;&nbsp;<input type=""checkbox"" name=""ARCHIVE"" value=""true"">Archived Posts" & vbNewLine)
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	'################# New Search Code #################################################
	Response.Write	"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Search By Date:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<select name=""SearchDate"">" & vbNewLine & _
					"<option value=""0"">Any Date</option>" & vbNewLine & _
					"<option value=""1"">Since Yesterday</option>" & vbNewLine & _
					"<option value=""2"">Since 2 Days Ago</option>" & vbNewLine & _
					"<option value=""5"">Since 5 Days Ago</option>" & vbNewLine & _
					"<option value=""7"">Since 7 Days Ago</option>" & vbNewLine & _
					"<option value=""14"">Since 14 Days Ago</option>" & vbNewLine & _
					"<option value=""30"">Since 30 Days Ago</option>" & vbNewLine & _
					"<option value=""60"">Since 60 Days Ago</option>" & vbNewLine & _
					"<option value=""120"">Since 120 Days Ago</option>" & vbNewLine & _
					"<option value=""365"">In the Last Year</option>" & vbNewLine & _
					"</select>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Search By Member:&nbsp;</td>" & vbNewLine
	if strUseMemberDropDownBox = 0 then
		Response.Write	"<td class=""formvalue""><input name=""SearchMember"" value="""" size=""25""></td>" & vbNewLine
	else
		Response.Write	"<td class=""formvalue"">" & vbNewLine & _
						"<select name=""SearchMember"">" & vbNewLine & _
						"<option value=""0"">All Members</option>" & vbNewLine
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
		strSql = strSql & " WHERE M_STATUS = " & 1
		strSql = strSql & " ORDER BY M_NAME ASC;"
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

		if rs.EOF then
			recMemberCount = ""
		else
			allMemberData = rs.GetRows(adGetRowsRest)
			recMemberCount = UBound(allMemberData,2)
			meMEMBER_ID = 0
			meM_NAME = 1
		end if

		rs.close
		set rs = nothing

		if recMemberCount <> "" then
			for iMember = 0 to recMemberCount
				MembersMemberID = allMemberData(meMEMBER_ID, iMember)
				MembersMemberName = allMemberData(meM_NAME, iMember)
				Response.Write	"<option value=""" & MembersMemberID & """>" & ChkString(MembersMemberName,"display") & "</option>" & vbNewline
			next
		end if
		Response.Write	"</select>" & vbNewLine & _
						"</td>" & vbNewLine
	end if
	Response.Write	"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"">Search</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if 
WriteFooter
Response.End
 
sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then
		mxpages = int(mxpages) + 1
	end if
	if mxpages > 1 then
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine)
		Response.Write("<tr>" & vbNewLine)
		Response.Write("<td valign=""bottom""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & getCurrentIcon(strIconPosticon,"","") & "</font></td>" & vbNewLine)
		for counter = 1 to mxpages
			ref = "<td align=""right"" valign=""bottom"" bgcolor=""" & strForumCellColor  & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<span class=""spnMessageText""><a href=""topic.asp?"
			ref = ref & ArchiveLink
            		ref = ref & "TOPIC_ID=" & Topic_ID
			ref = ref & "&whichpage=" & counter
			ref = ref & SearchLink
			ref = ref & """>" & counter & "</a></span></font></td>"
			Response.Write ref & vbNewLine
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr>" & vbNewLine)
				Response.Write("<tr>" & vbNewLine)
				Response.Write("<td>&nbsp;</td>" & vbNewLine)
			end if
		next				
	        Response.Write("</tr>" & vbNewLine)
	        Response.Write("</table>" & vbNewLine)
	end if
end sub

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write	"<form name=""PageNum" & fnum & """ action=""search.asp?" & chkString(Request.QueryString,"SQLString") & """ method=""post"">" & vbNewLine
		Response.Write	"<td>" & vbNewLine
		Response.Write	"<input name=""Search"" type=""hidden"" value=""" & trim(chkString(Request.Form("Search"),"search")) & """>" & vbNewLine
		Response.Write	"<input name=""andor"" type=""hidden"" value=""" & cLng(Request.Form("andor")) & """>" & vbNewLine
		Response.Write	"<input name=""Forum"" type=""hidden"" value=""" & cLng(Request.Form("Forum")) & """>" & vbNewLine
		Response.Write	"<input name=""SearchMessage"" type=""hidden"" value=""" & cLng(Request.Form("SearchMessage")) & """>" & vbNewLine
		if strArchiveState = "1" and ArchiveView = "true" then Response.Write("          <input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & vbNewLine)
		Response.Write	"<input name=""SearchDate"" type=""hidden"" value=""" & cLng(Request.Form("SearchDate")) & """>" & vbNewLine
		if strUseMemberDropDownBox = 0 then
			Response.Write	"<input name=""SearchMember"" type=""hidden"" value=""" & chkString(Request.Form("SearchMember"),"display") & """>" & vbNewLine
		else
			Response.Write	"<input name=""SearchMember"" type=""hidden"" value=""" & cLng(Request.Form("SearchMember")) & """>" & vbNewLine
		end if
		if fnum = 1 then
			Response.Write("<b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		else
			Response.Write("<b>There are " & maxpages & " Pages of Search Results: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then   
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			end if
		next
		if fnum = 1 then
			Response.Write("</select><b> of " & maxPages & "</b>" & vbNewLine)
		else
			Response.Write("</select>" & vbNewLine)
		end if
		Response.Write("</td>" & vbNewLine)
		Response.Write("</form>" & vbNewLine)
	end if
end sub

Function DoLastPostLink()
	if Topic_Replies < 1 or Topic_LastPostReplyID = 0 then
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & SearchLink & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","align=""absmiddle""") & "</a>"
	elseif Topic_LastPostReplyID <> 0 then
		PageLink = "whichpage=-1&"
		AnchorLink = "&REPLY_ID="
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & PageLink & "TOPIC_ID=" & Topic_ID & SearchLink & AnchorLink & Topic_LastPostReplyID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","align=""absmiddle""") & "</a>"
	else
		DoLastPostLink = ""
	end if
end function
%>
