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
'## Do Cookie stuffs with reload
nRefreshTime = Request.Cookies(strCookieURL & "Reload")

if Request.form("cookie") = "1" then
	if strSetCookieToForum = 1 then	
		Response.Cookies(strCookieURL & "Reload").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "Reload") = Request.Form("RefreshTime")
	Response.Cookies(strCookieURL & "Reload").expires = strForumTimeAdjust + 365
	nRefreshTime = Request.Form("RefreshTime")
end if

if nRefreshTime = "" then
	nRefreshTime = 0
end if

ActiveSince = Request.Cookies(strCookieURL & "ActiveSince")
'## Do Cookie stuffs with show last date
if Request.form("cookie") = "2" then
	ActiveSince = Request.Form("ShowSinceDateTime")
	if strSetCookieToForum = 1 then	
      		Response.Cookies(strCookieURL & "ActiveSince").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "ActiveSince") = ActiveSince
end if
Dim ModerateAllowed
Dim HasHigherSub
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts
Dim canView
HasHigherSub = false
%>
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<%
Select Case ActiveSince
	Case "LastVisit"
		lastDate = ""
	Case "LastFifteen"
 		lastDate = DateToStr(DateAdd("n",-15,strForumTimeAdjust))
	Case "LastThirty"
 		lastDate = DateToStr(DateAdd("n",-30,strForumTimeAdjust))
	Case "LastFortyFive"
 		lastDate = DateToStr(DateAdd("n",-45,strForumTimeAdjust))
	Case "LastHour"
		lastDate = DateToStr(DateAdd("h",-1,strForumTimeAdjust))
	Case "Last2Hours"
		lastDate = DateToStr(DateAdd("h",-2,strForumTimeAdjust))
	Case "Last6Hours"
		lastDate = DateToStr(DateAdd("h",-6,strForumTimeAdjust))
	Case "Last12Hours"
		lastDate = DateToStr(DateAdd("h",-12,strForumTimeAdjust))
	Case "LastDay"
		lastDate = DateToStr(DateAdd("d",-1,strForumTimeAdjust))
	Case "Last2Days"
		lastDate = DateToStr(DateAdd("d",-2,strForumTimeAdjust))
	Case "LastWeek"
		lastDate = DateToStr(DateAdd("ww",-1,strForumTimeAdjust))
	Case "Last2Weeks"
		lastDate = DateToStr(DateAdd("ww",-2,strForumTimeAdjust))
	Case "LastMonth"
		lastDate = DateToStr(DateAdd("m",-1,strForumTimeAdjust))
	Case "Last2Months"
		lastDate = DateToStr(DateAdd("m",-2,strForumTimeAdjust))
	Case Else
		lastDate = ""
End Select

Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"<!--" & vbNewLine & _
		"function autoReload()	{ 	document.ReloadFrm.submit()		}" & vbNewLine & _
		"function SetLastDate()	{	document.LastDateFrm.submit()	}" & vbNewLine & _
		"function jumpTo(s)	{	if (s.selectedIndex != 0) location.href = s.options[s.selectedIndex].value;return 1;}" & vbNewLine & _
		"//defaultStatus = ""You last loaded this page on " & chkDate(DateToStr(strForumTimeAdjust)," ",true) & " (Forum Time)""" & vbNewLine & _
		"// -->" & vbNewLine & _
		"</script>" & vbNewLine

if IsEmpty(Session(strCookieURL & "last_here_date")) then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTUserName)
end if
if lastDate = "" then
	lastDate = Session(strCookieURL & "last_here_date")
end if
if Request.Form("AllRead") = "Y" then
	lastDate = ChkString(Request.Form("BuildTime"),"SQLString")
	'## The redundant line below is necessary, don't delete it.
	Session(strCookieURL & "last_here_date") = lastDate
	Session(strCookieURL & "last_here_date") = lastDate
	UpdateLastHereDate lastDate,strDBNTUserName
	ActiveSince = ""
end if

if strModeration = "1" and mLev > 2 then
	UnModeratedPosts = CheckForUnmoderatedPosts("BOARD", 0, 0, 0)
end if

' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs, strTopicSubs
If MySubCount > 0 then
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
End If

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

'## Forum_SQL - Get all active topics from last visit
strSql = "SELECT F.FORUM_ID, " & _
         "F.F_SUBJECT, " & _
	 "F.F_SUBSCRIPTION, " & _
	 "F.F_STATUS, " & _
         "C.CAT_ID, " & _
	 "C.CAT_NAME, " & _
	 "C.CAT_SUBSCRIPTION, " & _
	 "C.CAT_STATUS, " & _
	 "T.T_STATUS, " & _
	 "T.T_VIEW_COUNT, " & _
	 "T.TOPIC_ID, " & _
	 "T.T_SUBJECT, " & _
	 "T.T_AUTHOR, " & _
	 "T.T_REPLIES, " & _
	 "T.T_UREPLIES, " & _
	 "M.M_NAME, " & _
	 "T.T_LAST_POST_AUTHOR, " & _
	 "T.T_LAST_POST, " & _
	 "T.T_LAST_POST_REPLY_ID, " & _
	 "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME, " & _
	 "F.F_PRIVATEFORUMS, " & _
	 "F.F_PASSWORD_NEW " & _
	 "FROM " & strMemberTablePrefix & "MEMBERS M, " & _
	 strTablePrefix & "FORUM F, " & _
	 strTablePrefix & "TOPICS T, " & _
	 strTablePrefix & "CATEGORY C, " & _
	 strMemberTablePrefix & "MEMBERS MEMBERS_1 " & _
	 "WHERE T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID "
if strPrivateForums = "1" and mLev < 4 then
	strSql = strSql & " AND F.FORUM_ID IN (" & allAllowedForums & ") "
end if
Dim strDeniedForums : strDeniedForums = Session(strCookieURL & "DeniedForums" & MemberID)
if strDeniedForums <> "" and strDeniedForums <> "-1" then
	strSql = strSql & " AND F.FORUM_ID NOT IN(" & strDeniedForums & ") "
end if
strSql = strSql & "AND F.F_TYPE = 0 " & _
	 "AND F.FORUM_ID = T.FORUM_ID " & _
	 "AND C.CAT_ID = T.CAT_ID " & _
	 "AND M.MEMBER_ID = T.T_AUTHOR " & _
	 "AND (T.T_LAST_POST > '" & lastDate & "'"

' DEM --> if not an admin, all unapproved posts should not be viewed.
if mlev <> 4 then
	strSql = strSql & " AND ((T.T_AUTHOR <> " & MemberID &_
			  " AND T.T_STATUS < 2)"  ' Ignore unapproved/held posts
	if mlev = 3 and ModOfForums <> "" then
		strSql = strSql & " OR T.FORUM_ID IN (" & ModOfForums & ") "
	end if
	strSql = strSql & "  OR T.T_AUTHOR = " & MemberID & ")"
end if
if Group > 1 and strGroupCategories = "1" then
	strSql = strSql & " AND (C.CAT_ID = 0"
	if recGroupCatCount <> "" then
		for iGroupCat = 0 to recGroupCatCount
			strSql = strSql & " or C.CAT_ID = " & allGroupCatData(1, iGroupCat)
		next
		strSql = strSql & ")"
	else
		strSql = strSql & ")"
	end if
end if

strSql = strSql & ") "
strSql = strSql & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT, T.T_LAST_POST DESC "

Set rs = Server.CreateObject("ADODB.Recordset")
if strDBType <> "mysql" then rs.cachesize = 50
rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rs.EOF then
	recActiveTopicsCount = ""
else
	allActiveTopics = rs.GetRows(adGetRowsRest)
	recActiveTopicsCount = UBound(allActiveTopics,2)
end if

rs.close
set rs = nothing

' Sets up the Tree structure at the top of the page
Response.Write	"<table class=""misc"" width=""100%"" cellspacing=""0"">" & vbNewline & _
				"<tr>" & vbNewline & _
				"<form name=""LastDateFrm"" action=""active.asp"" method=""post"">" & vbNewline & _
				"<td class=""secondnav"">" & vbNewline & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;" & _
				"Active Topics Since " & vbNewLine
Response.Write	"<select name=""ShowSinceDateTime"" size=""1"" onchange=""SetLastDate();"">" & vbNewline & _
		"<option value=""LastVisit"""
if ActiveSince = "LastVisit" or ActiveSince = "" then
	Response.Write " selected"
end if
Response.Write	">&nbsp;Last Visit on " & ChkDate(Session(strCookieURL & "last_here_date"),"",true) & "&nbsp;</option>" & vbNewline & _
				"<option value=""LastFifteen""" & chkSelect(ActiveSince,"LastFifteen") & ">&nbsp;Last 15 minutes</option>" & vbNewline & _
				"<option value=""LastThirty""" & chkSelect(ActiveSince,"LastThirty") & ">&nbsp;Last 30 minutes</option>" & vbNewline & _
				"<option value=""LastFortyFive""" & chkSelect(ActiveSince,"LastFortyFive") & ">&nbsp;Last 45 minutes</option>" & vbNewline & _
				"<option value=""LastHour""" & chkSelect(ActiveSince,"LastHour") & ">&nbsp;Last Hour</option>" & vbNewline & _
				"<option value=""Last2Hours""" & chkSelect(ActiveSince,"Last2Hours") & ">&nbsp;Last 2 Hours</option>" & vbNewline & _
				"<option value=""Last6Hours""" & chkSelect(ActiveSince,"Last6Hours") & ">&nbsp;Last 6 Hours</option>" & vbNewline & _
				"<option value=""Last12Hours""" & chkSelect(ActiveSince,"Last12Hours") & ">&nbsp;Last 12 Hours</option>" & vbNewline & _
				"<option value=""LastDay""" & chkSelect(ActiveSince,"LastDay") & ">&nbsp;Yesterday</option>" & vbNewline & _
				"<option value=""Last2Days""" & chkSelect(ActiveSince,"Last2Days") & ">&nbsp;Last 2 Days</option>" & vbNewline & _
				"<option value=""LastWeek""" & chkSelect(ActiveSince,"LastWeek") & ">&nbsp;Last Week</option>" & vbNewline & _
				"<option value=""Last2Weeks""" & chkSelect(ActiveSince,"Last2Weeks") & ">&nbsp;Last 2 Weeks</option>" & vbNewline & _
				"<option value=""LastMonth""" & chkSelect(ActiveSince,"LastMonth") & ">&nbsp;Last Month</option>" & vbNewline & _
				"<option value=""Last2Months""" & chkSelect(ActiveSince,"Last2Months") & ">&nbsp;Last 2 Months</option>" & vbNewline & _
				"</select>" & vbNewline

Response.Write	"<input type=""hidden"" name=""Cookie"" value=""2"">" & vbNewLine & _
				"</td>" & vbNewline & _
				"</form>" & vbNewline & _
				"<form name=""ReloadFrm"" action=""active.asp"" method=""post"">" & vbNewline & _
				"<td class=""reloadmenu"">" & vbNewline & _
				"<select name=""RefreshTime"" size=""1"" onchange=""autoReload();"">" & vbNewline & _
				"<option value=""0""" & chkSelect(nRefreshTime,0) & ">Don't reload automatically</option>" & vbNewline & _
				"<option value=""1""" & chkSelect(nRefreshTime,1) & ">Reload page every minute</option>" & vbNewline & _
				"<option value=""2""" & chkSelect(nRefreshTime,2) & ">Reload page every 2 minutes</option>" & vbNewline & _
				"<option value=""5""" & chkSelect(nRefreshTime,5) & ">Reload page every 5 minutes</option>" & vbNewline & _
				"<option value=""10""" & chkSelect(nRefreshTime,10) & ">Reload page every 10 minutes</option>" & vbNewline & _
				"<option value=""15""" & chkSelect(nRefreshTime,15) & ">Reload page every 15 minutes</option>" & vbNewline & _
				"</select>" & vbNewline
Response.Write	"<input type=""hidden"" name=""Cookie"" value=""1"">" & vbNewline & _
				"</td>" & vbNewline & _
				"</form>" & vbNewline & _
				"</tr>" & vbNewline & _
				"</table>" & vbNewline

'### Start to build the table
Response.Write	"<table class=""content"" width=""100%"" cellspacing=""0"" cellpadding=""4"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewline & _
				"<td class=""statusicon"">" & vbNewline
If recActiveTopicsCount <> "" and (mLev > 0) then
	Response.Write	"<form name=""MarkRead"" action=""active.asp"" method=""post"" style=""display:inline"">" & vbNewline & _
					"<input type=""hidden"" name=""AllRead"" value=""Y"">" & vbNewline & _
					"<input type=""hidden"" name=""BuildTime"" value=""" & DateToStr(strForumTimeAdjust) & """>" & vbNewline & _
					"<input type=""hidden"" name=""Cookie"" value=""2"">" & vbNewLine & _
					"<acronym title=""Mark all topics as read""><input type=""image"" class=""image"" src=""" & strImageUrl & "icon_topic_all_read.gif"" value=""Mark all read"" id=""submit1"" name=""Mark all topics as read""" & dWStatus("Mark all topics as read") & "></acronym>" & vbNewLine & _
					"</form>" & vbNewline
else 
	Response.Write	"&nbsp;" & vbNewline
end if
Response.Write	"</td>" & vbNewLine
Response.Write	"<td>Topic</td>" & vbNewline & _
				"<td>Author</td>" & vbNewline & _
				"<td>Replies</td>" & vbNewline & _
				"<td>Read</td>" & vbNewline & _
				"<td>Last Post</td>" & vbNewline
if (mlev > 0) or (lcase(strNoCookies) = "1") then
	Response.Write	"<td class=""options"">"
	if (mLev = 4 or mLev = 3) or (lcase(strNoCookies) = "1") then
        	if UnModeratedPosts > 0 then
        		UnModeratedFPosts = 0
			Response.Write "<a href=""moderate.asp"">" & getCurrentIcon(strIconFolderModerate,"View All UnModerated Posts","hspace=""0""") & "</a>"
		else
			Response.Write("&nbsp;")
	        end if
	else
		Response.Write("&nbsp;")
	end if
	Response.Write	"</td>" & vbNewline
end if
Response.Write	"</tr>" & vbNewline
if recActiveTopicsCount = "" then
	Response.Write	"<tr>" & vbNewline & _
					"<td colspan=""7""><b>No Active Topics Found</b></td>" & vbNewline & _
					"</tr>" & vbNewline
else
	currForum = 0
	fDisplayCount = 0
	canAccess = 0 	
	
	fFORUM_ID = 0
	fF_SUBJECT = 1
	fF_SUBSCRIPTION = 2
	fF_STATUS = 3
	fCAT_ID = 4
	fCAT_NAME = 5
	fCAT_SUBSCRIPTION = 6
	fCAT_STATUS = 7
	fT_STATUS = 8
	fT_VIEW_COUNT = 9
	fTOPIC_ID = 10
	fT_SUBJECT = 11
	fT_AUTHOR = 12
	fT_REPLIES = 13
	fT_UREPLIES = 14
	fM_NAME = 15
	fT_LAST_POST_AUTHOR = 16
	fT_LAST_POST = 17
	fT_LAST_POST_REPLY_ID = 18
	fLAST_POST_AUTHOR_NAME = 19
	fF_PRIVATEFORUMS = 20
	fF_PASSWORD_NEW = 21

	for RowCount = 0 to recActiveTopicsCount
		'## Store all the recordvalues in variables first.

		Forum_ID = allActiveTopics(fFORUM_ID,RowCount)
		Forum_Subject = allActiveTopics(fF_SUBJECT,RowCount)
		ForumSubscription = allActiveTopics(fF_SUBSCRIPTION,RowCount)
		Forum_Status = allActiveTopics(fF_STATUS,RowCount)
		Cat_ID = allActiveTopics(fCAT_ID,RowCount)
		Cat_Name = allActiveTopics(fCAT_NAME,RowCount)
		CatSubscription = allActiveTopics(fCAT_SUBSCRIPTION,RowCount)
		Cat_Status = allActiveTopics(fCAT_STATUS,RowCount)
		Topic_Status = allActiveTopics(fT_STATUS,RowCount)
		Topic_View_Count = allActiveTopics(fT_VIEW_COUNT,RowCount)
		Topic_ID = allActiveTopics(fTOPIC_ID,RowCount)
		Topic_Subject = allActiveTopics(fT_SUBJECT,RowCount)
		Topic_Author = allActiveTopics(fT_AUTHOR,RowCount)
		Topic_Replies = allActiveTopics(fT_REPLIES,RowCount)
		Topic_UReplies = allActiveTopics(fT_UREPLIES,RowCount)
		Member_Name = allActiveTopics(fM_NAME,RowCount)
		Topic_Last_Post_Author = allActiveTopics(fT_LAST_POST_AUTHOR,RowCount)
		Topic_Last_Post = allActiveTopics(fT_LAST_POST,RowCount)
		Topic_Last_Post_Reply_ID = allActiveTopics(fT_LAST_POST_REPLY_ID,RowCount)
		Topic_Last_Post_Author_Name = chkString(allActiveTopics(fLAST_POST_AUTHOR_NAME,RowCount),"display")
		Forum_PrivateForums = allActiveTopics(fF_PRIVATEFORUMS,RowCount)
		Forum_FPasswordNew = allActiveTopics(fF_PASSWORD_NEW,RowCount)
		
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
		fDisplayCount = fDisplayCount + 1
		' -- Display forum name
		if currForum <> Forum_ID then
			Response.Write	"<tr class=""section"">" & vbNewline & _
					"<td colspan=""6"">" & _
					"<a href=""default.asp?CAT_ID=" & Cat_ID & """><b>" & ChkString(Cat_Name,"display") & "</b></a>&nbsp;/&nbsp;" & _
					"<a href=""forum.asp?FORUM_ID=" & Forum_ID & """><b>" & ChkString(Forum_Subject,"display") & "</b></a></td>" & vbNewline
			if (mlev > 0) or (lcase(strNoCookies) = "1") then 
				Response.Write	"<td class=""options"">" & vbNewLine
				if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
					ForumAdminOptions
				else
					if Cat_Status <> 0 and Forum_Status <> 0 then
						ForumMemberOptions
					else
						Response.Write "&nbsp;" & vbNewLine
					end if
				end if 
				Response.Write	"</td>" & vbNewline
			elseif (mLev = 3) then
				Response.Write	"<td>&nbsp;</td>" & vbNewline
			end if
			Response.Write	"</tr>" & vbNewline
		end if
		Response.Write	"<tr>" & vbNewline
		Response.Write	"<td class=""statusicon"">"
		' -- Set up a link to the topic and display the icon appropriate to the status of the post.
		Response.Write "<a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>"
		' - If status = 0, topic/forum/category is locked.  If status > 2, posts are unmoderated/rejected
		if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
			' DEM --> Added code for topic moderation
			if Topic_Status = 2 then
				UnApprovedFound = "Y"
				Response.Write 	getCurrentIcon(strIconFolderUnmoderated,"Topic Not Moderated","hspace=""0""") & "</a>" & vbNewline
			elseif Topic_Status = 3 then
				HeldFound = "Y"
				Response.Write 	getCurrentIcon(strIconFolderHold,"Topic on Hold","hspace=""0""") & "</a>" & vbNewline
				' DEM --> end of code Added for topic moderation
			elseif lcase(strHotTopic) = "1" and Topic_Replies >= intHotTopicNum then
				Response.Write	getCurrentIcon(strIconFolderNewHot,"Hot Topic with New Posts","hspace=""0""") & "</a>" & vbNewline
			elseif Topic_Last_Post < lastdate then
				Response.Write	getCurrentIcon(strIconFolder,"No New Posts","") & "</a>" & vbNewline
			else
				Response.Write	getCurrentIcon(strIconFolderNew,"New Posts","") & "</a>" & vbNewline
			end if
		else
			if Cat_Status = 0 then
				strAltText = "Category locked"
			elseif Forum_Status = 0 then
				strAltText = "Forum locked"
			else
				strAltText = "Topic locked"
			end if
			if Topic_Last_Post < lastdate then
				Response.Write	getCurrentIcon(strIconFolderLocked,strAltText,"hspace=""0""")
			else
				Response.Write	getCurrentIcon(strIconFolderNewLocked,strAltText,"hspace=""0""")
			end if
			Response.Write	"</a>" & vbNewline
		end if
		Response.Write	"</td>" & vbNewline
		Response.Write	"<td>"
		Response.Write	"<div class=""forumlink""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & ChkString(Topic_Subject,"title") & "</a></div>" & vbNewline
		if strShowPaging = "1" then
			TopicPaging()
		end if
		Response.Write	"</td>" & vbNewline
		Response.Write	"<td>" & profileLink(chkString(Member_Name,"display"),Topic_Author) & "</td>" & vbNewline
		Response.Write	"<td class=""counts"">" & Topic_Replies & "</td>" & vbNewline
		Response.Write	"<td class=""counts"">" & Topic_View_Count & "</td>" & vbNewline
		if IsNull(Topic_Last_Post_Author) then
			strLastAuthor = ""
		else
			strLastAuthor = "<br />by:&nbsp;" & profileLink(Topic_Last_Post_Author_Name,Topic_Last_Post_Author)
			if strJumpLastPost = "1" then strLastAuthor = strLastAuthor & "&nbsp;" & DoLastPostLink
		end if
		Response.Write	"<td class=""counts""><span class=""lastpost"">" & ChkDate(Topic_Last_Post, "&nbsp;" ,true) & _
						"</span><span class=""lastuser"">" & strLastAuthor & "</span></td>" & vbNewline
		if (mlev > 0) or (lcase(strNoCookies) = "1") then
			Response.Write	"<td class=""options""><b>" & vbNewLine
			if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
				call TopicAdminOptions
			else
				if  Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
					call TopicMemberOptions
				else
					Response.Write	"&nbsp;" & vbNewline
				end if
			end if
			Response.Write	"</b></td>" & vbNewline
		elseif (mLev = 3) then
			Response.Write	"<td>&nbsp;</td>" & vbNewline
		end if
		Response.Write	"</tr>" & vbNewline
		currForum = Forum_ID
	next
	if fDisplayCount = 0 then
		Response.Write	"<tr>" & vbNewline & _
						"<td colspan=""" & aGetColspan(7,6) & """><b>No Active Topics Found</b></td>" & vbNewline & _
						"</tr>" & vbNewline
	end if
end if
Response.Write	"</table>" & vbNewline

Response.Write	"<table class=""misc2"" width=""100%"">" & vbNewline & _
				"<tr>" & vbNewline & _
				"<td class=""legend"">" & vbNewline & _
				getCurrentIcon(strIconFolderNew,"New Posts","align=""absmiddle""") & " New posts since last logon.<br />" & vbNewLine & _
				getCurrentIcon(strIconFolder,"Old Posts","align=""absmiddle""") & " Old Posts."
if lcase(strHotTopic) = "1" then Response.Write	(" (" & getCurrentIcon(strIconFolderHot,"Hot Topic","align=""absmiddle""") & "&nbsp;" & intHotTopicNum & " replies or more.)" & vbNewLine)
Response.Write	"<br />" & getCurrentIcon(strIconFolderLocked,"Locked Topic","align=""absmiddle""") & " Locked topic." & vbNewLine
' DEM --> Start of Code added for moderation
if HeldFound = "Y" then
	Response.Write "<br />" & getCurrentIcon(strIconFolderHold,"Held Topic","align=""absmiddle""") & " Held Topic." & vbNewline
end if
if UnapprovedFound = "Y" then
	Response.Write "<br />" & getCurrentIcon(strIconFolderUnmoderated,"UnModerated Topic","align=""absmiddle""") & " UnModerated Topic." & vbNewline
end if
' DEM --> End of Code added for moderation
Response.Write	"</td>" & vbNewLine & _
				"<td class=""jumpmenu"">" & vbNewline
%>
        <!--#INCLUDE FILE="inc_jump_to.asp" -->
<%
Response.Write 	"</td>" & vbNewline & _
		"</tr>" & vbNewline & _
		"</table>" & vbNewline & _
		"<script language=""javascript"" type=""text/javascript"">" & vbNewline & _
		"<!--" & vbNewline & _
		"if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {" & vbNewline & _
		"reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value" & vbNewline & _
		"self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)" & vbNewline & _
		"}" & vbNewline & _
		"//-->" & vbNewline & _
		"</script>" & vbNewline
WriteFooter
Response.End

sub ForumAdminOptions()
	if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
		if Cat_Status = 0 then
			if mlev = 4 then
				Response.Write	"<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Category","") & "</a>" & vbNewline
			else
				Response.Write	getCurrentIcon(strIconFolderLocked,"Category Locked","") & vbNewline
			end if
		else
			if Forum_Status <> 0 then
				Response.Write	"<a href=""JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderLocked,"Lock Forum","") & "</a>" & vbNewline
			else
				Response.Write	"<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Forum","") & "</a>" & vbNewline
			end if
		end if
		if (Cat_Status <> 0 and Forum_Status <> 0) or (ModerateAllowed = "Y") then
			Response.Write	"<a href=""post.asp?method=EditForum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "&type=0"">" & getCurrentIcon(strIconFolderPencil,"Edit Forum Properties","hspace=""0""") & "</a>" & vbNewline
		end if
		if mLev = 4 or lcase(strNoCookies) = "1" then Response.Write("<a href=""JavaScript:openWindow('pop_delete.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderDelete,"Delete Forum","") & "</a>" & vbNewLine)
		Response.Write	"<a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>" & vbNewLine
 		' DEM --> Start of Code added to handle subscription processing.
		if (strSubscription < 4 and strSubscription > 0) and (CatSubscription > 0) and ForumSubscription = 1 and strEmail = 1 then
			if InArray(strForumSubs, Forum_ID) then
				Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, 0, "N")
			elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,Cat_ID)) then
				Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, 0, "N")
			end if
		end if
		' DEM --> End of code added to handle subscription processing.
	end if
end sub

sub ForumMemberOptions()
	if (mlev > 0) then
		Response.Write	"<a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>" & vbNewLine
 		' DEM --> Start of Code added to handle subscription processing.
	        if (strSubscription > 0 and strSubscription < 4) and CatSubscription > 0 and ForumSubscription = 1 and strEmail = 1 then
				if InArray(strForumSubs, Forum_ID) then
					Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, 0, "N")
				elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,Cat_ID)) then
					Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, 0, "N")
				end if
        	end if
	end if
end sub

sub TopicAdminOptions()
	if Cat_Status = 0 then
		Response.Write	"<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Category","hspace=""0""") & "</a>" & vbNewLine
	elseif Forum_Status = 0 then
		Response.Write	"<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Forum","hspace=""0""") & "</a>" & vbNewLine
	elseif Topic_Status <> 0 then
		Response.Write	"<a href=""JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconLock,"Lock Topic","hspace=""0""") & "</a>" & vbNewLine
	else
		Response.Write	"<a href=""JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	if (ModerateAllowed = "Y") or (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) then
		Response.Write	"<a href=""post.asp?method=EditTopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&auth=" & Topic_Author & """>" & getCurrentIcon(strIconPencil,"Edit Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	Response.Write	"<a href=""JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","hspace=""0""") & "</a>" & vbNewLine
	if Topic_Status <= 1 then
		Response.Write	"<a href=""post.asp?method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	' DEM --> Start of Code for Full Moderation
        if Topic_Status > 1 then
		TopicString = "TOPIC_ID=" & Topic_ID & "&CAT_ID=" & Cat_ID & "&FORUM_ID=" & Forum_ID
               	Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Topic","hspace=""0""") & "</a>" & vbNewline
        end if
	' DEM --> End of Code for Full Moderation 
	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription < 4 and strSubscription > 0) and (CatSubscription > 0) and ForumSubscription > 0 and strEmail = 1 then
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write "&nbsp;" & ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N")
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write "&nbsp;" & ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N")
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub

sub TopicMemberOptions()
        if (Topic_Status > 0 and Topic_Author = MemberID) or (ModerateAllowed = "Y") then
		Response.Write	"<a href=""post.asp?method=EditTopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconPencil,"Edit Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
        if (Topic_Status > 0 and Topic_Author = MemberID and Topic_Replies = 0) or (ModerateAllowed = "Y") then
		Response.Write	"<a href=""JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	if Topic_Status <= 1 then
		Response.Write	"<a href=""post.asp?method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	if (strSubscription < 4 and strSubscription > 0) and (CatSubscription > 0) and ForumSubscription > 0 and strEmail = 1 then
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write "&nbsp;" & ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N")
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write "&nbsp;" & ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N")
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub

sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then
		mxpages = int(mxpages) + 1
	end if
	if mxpages > 1 then
		Response.Write	"<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"<tr class=""ffs"">" & vbNewLine & _
				"<td valign=""bottom"">" & getCurrentIcon(strIconPosticon,"","align=""absmiddle""") & "</td>" & vbNewLine
		for counter = 1 to mxpages
			ref =	"<td align=""right"" valign=""bottom"">"
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if
			ref = ref & widenum(counter) & "<a href=""topic.asp?"
			ref = ref & ArchiveLink
			ref = ref & "TOPIC_ID=" & Topic_ID
			ref = ref & "&whichpage=" & counter
			ref = ref & """>" & counter & "</a></td>"
			Response.Write ref & vbNewLine
			if counter mod strPageNumberSize = 0 and counter < mxpages then
				Response.Write("</tr>" & vbNewLine)
				Response.Write("<tr>" & vbNewLine)
				Response.Write("<td>&nbsp;</td>" & vbNewLine)
			end if
		next
		Response.Write("</tr>" & vbNewLine)
		Response.Write("</table>" & vbNewLine)
	end if
end sub

Function DoLastPostLink()
	if Topic_Replies < 1 or Topic_Last_Post_Reply_ID = 0 then
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","align=""absmiddle""") & "</a>"
	elseif Topic_Last_Post_Reply_ID <> 0 then
		PageLink = "whichpage=-1&"
		AnchorLink = "&REPLY_ID="
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & PageLink & "TOPIC_ID=" & Topic_ID & AnchorLink & Topic_Last_Post_Reply_ID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","align=""absmiddle""") & "</a>"
	else
		DoLastPostLink = ""
	end if
end function

function aGetColspan(lIN, lOUT)
	if (mlev > 0 or strNoCookies = "1") then lOut = lOut + 1
	if lOut > lIn then
		aGetColspan = lIN
	else
		aGetColspan = lOUT
	end if
end function
%>
