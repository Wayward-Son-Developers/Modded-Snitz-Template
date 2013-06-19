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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_func_common.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<%

dim intResults,Topic_ID,strSubject,Topic_Replies,Reply_ID


'Wed, 02 Oct 2002 13:00:00 0000 - RFC-822 date format

intResults = 25
TimeZoneOffset="-0500" 

strSql = "SELECT TOP " & intResults
strSql = strSql & " TopicId, ReplyId, Author, PageCounter, Title, Descript, PostDate, Category "
strSql = strSql & "FROM "
strSql = strSql & "(SELECT TOP 100 PERCENT "
strSql = strSql & "C.CAT_NAME AS Category, "
strSql = strSql & "R.TOPIC_ID AS TopicId, "
strSql = strSql & "R.REPLY_ID AS ReplyId, "
strSql = strSql & "M.M_NAME AS Author, "
strSql = strSql & "T.T_REPLIES AS PageCounter, "
strSql = strSql & "T.T_SUBJECT AS Title, "
strSql = strSql & "R.R_MESSAGE AS Descript, "
Select Case strDBType
	case "sqlserver", "mysql"
		strSql = strSql & "COALESCE(R.R_LAST_EDIT, R.R_DATE) AS PostDate "
	case "access"
		strSql = strSql & "IIf(R.R_LAST_EDIT, R.R_DATE) AS PostDate "
End Select
strSql = strSql & "FROM (((" & strTablePrefix & "REPLY R "
strSql = strSql & "INNER JOIN " & strTablePrefix & "TOPICS T ON R.TOPIC_ID = T.TOPIC_ID) "
strSql = strSql & "INNER JOIN " & strTablePrefix & "FORUM F ON R.FORUM_ID = F.FORUM_ID) "
strSql = strSql & "INNER JOIN " & strMemberTablePrefix & "MEMBERS M ON R.R_AUTHOR = M.MEMBER_ID) "
strSql = strSql & "INNER JOIN " & strTablePrefix & "CATEGORY C ON R.CAT_ID = C.CAT_ID "
strSql = strSql & "WHERE (F.F_PRIVATEFORUMS = 0) AND T.T_STATUS=1 "
strSql = strSql & "UNION ALL "
strSql = strSql & "SELECT TOP 100 PERCENT    "
strSql = strSql & "C.CAT_NAME AS Category, "
strSql = strSql & "T.TOPIC_ID AS TopicId, "
strSql = strSql & "(-1) AS ReplyId, "
strSql = strSql & "M.M_NAME AS Author, "
strSql = strSql & "T.T_REPLIES AS PageCounter, "
strSql = strSql & "T.T_SUBJECT AS Title, "
strSql = strSql & "T.T_MESSAGE AS Descript,"
strSql = strSql & "T.T_DATE AS PostDate "
strSql = strSql & "FROM ((" & strTablePrefix & "TOPICS T "
strSql = strSql & "INNER JOIN " & strTablePrefix & "FORUM F ON T.FORUM_ID = F.FORUM_ID) "
strSql = strSql & "INNER JOIN " & strMemberTablePrefix & "MEMBERS M ON T.T_AUTHOR = M.MEMBER_ID) "
strSql = strSql & "INNER JOIN " & strTablePrefix & "CATEGORY C ON T.CAT_ID = C.CAT_ID "
strSql = strSql & "WHERE (F.F_PRIVATEFORUMS = 0) AND T.T_STATUS=1 "
strSql = strSql & "ORDER BY PostDate DESC "
strSql = strSql & ") AS Posts "
strSql = strSql & "ORDER BY PostDate DESC;"

set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rs.EOF then
	recActiveTopicsCount = ""
else
	allActiveTopics = rs.GetRows(adGetRowsRest)
	recActiveTopicsCount = UBound(allActiveTopics,2)
end if

rs.close
set rs = nothing

xml = ""
xml = "<?xml version=""1.0"" encoding=""ISO-8859-1"" ?><!-- RSS generation done by Snitz Forums 2000 on " & chkDate(datetostr(strForumTimeAdjust)," ",true) & " --><rss version=""2.0""><channel>"
xml = xml & "<title>Latest Posts on " & strForumTitle & "</title>"
xml = xml & "<link>" & strForumURL & "</link>"
xml = xml & "<description>" & strForumTitle & "</description>"
xml = xml & "<image>"
xml = xml & "<link>" & strHomeURL & "</link>"
xml = xml & "<url>" & strForumURL & strImageURL & strTitleImage & "</url>"
'xml = xml & "<url>" & strForumURL & strImageURL & "logo_powered_by.gif</url>"
xml = xml & "<title>" & strForumTitle & " RSS Feed</title>"
xml = xml & "<width>142</width>"
xml = xml & "<height>23</height>"
xml = xml & "</image>"
if recActiveTopicsCount <> "" then
	TopicId = 0
	ReplyId = 1
	Author = 2
	Replies = 3
	Title = 4
	Description = 5
	PostDate = 6
	CatTitle = 7

	for RowCount = 0 to recActiveTopicsCount
		Topic_Replies = allActiveTopics(Replies,RowCount)
		Post_Subject = chkstring(replace(allActiveTopics(Title,RowCount),"&","&amp;"),"display")
		Post_Message = replace(allActiveTopics(Description,RowCount),"&","&amp;")
		Topic_ID = allActiveTopics(TopicId,RowCount)
		Reply_ID = allActiveTopics(ReplyId,RowCount)
		Post_Author = allActiveTopics(Author,RowCount)
		Post_Date = allActiveTopics(PostDate,RowCount)
		CategoryTitle = allActiveTopics(CatTitle,RowCount)

		xml = xml & "<item>"
		xml = xml & "<category><![CDATA[" & formatStr(CategoryTitle) & "]]></category>"
		xml = xml & "<title>" & Post_Subject & "</title>"
		xml = xml & "<author>" & Post_Author & "</author>"
		xml = xml & "<link>" & strForumURL & DoLastPostLink & "</link>"
		xml = xml & "<description><![CDATA[" & HTMLDecode(formatStr(Post_Message)) & "]]></description>"
		xml = xml & "<pubDate>" & RFC822_Date(Post_Date,TimeZoneOffset) & "</pubDate>"
		xml = xml & "</item>"
	next
end if
xml = xml & "</channel></rss>"
Response.Clear
Response.Expires = 0
Response.ContentType = "text/xml"
Response.Write xml

my_Conn.close
set my_Conn = nothing
Response.End

Function DoLastPostLink()
	if Reply_ID = -1 then
		DoLastPostLink = "topic.asp?TOPIC_ID=" & Topic_ID
	elseif Reply_ID > 0 then
		PageLink = "whichpage=-1&amp;"
		AnchorLink = "&amp;REPLY_ID="
		DoLastPostLink = "topic.asp?" & PageLink & "TOPIC_ID=" & Topic_ID & AnchorLink & Reply_ID
	else
		DoLastPostLink = "topic.asp?TOPIC_ID=" & Topic_ID
	end if
end function

Function RFC822_Date(myDate, offset)
   Dim myDay, myDays, myMonth, myYear
   Dim myHours, myMonths, mySeconds

   myDate = strToDate(myDate)
   myDay = WeekdayName(Weekday(myDate),true)
   myDays = Day(myDate)
   myMonth = MonthName(Month(myDate), true)
   myYear = Year(myDate)
   myHours = zeroPad(Hour(myDate), 2)
   myMinutes = zeroPad(Minute(myDate), 2)
   mySeconds = zeroPad(Second(myDate), 2)

   RFC822_Date = myDay&", "& _
                                  myDays&" "& _
                                  myMonth&" "& _ 
                                  myYear&" "& _
                                  myHours&":"& _
                                  myMinutes&":"& _
                                  mySeconds&" "& _ 
                                  offset
End Function 
Function zeroPad(m, t)
   zeroPad = String(t-Len(m),"0")&m
End Function


%>