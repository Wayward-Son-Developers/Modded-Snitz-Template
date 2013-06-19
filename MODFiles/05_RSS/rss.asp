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
set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

dim intResults,Topic_ID,strSubject,Topic_Replies,Topic_Last_Post_Reply_ID
intResults = 20
'disable images
strIcons = "1"
strIMGInPosts = "1"

if Request.QueryString("MEMBER_ID")<>"" then
	submittedUsrID = cLng(Request.QueryString("MEMBER_ID"))
	submittedChkID = Request.QueryString("ChkID")
	
	if chkRssUsr(submittedUsrID,submittedChkID) > 0 then
		strTitleOwner = " for " & getMemberName(submittedUsrID)
		
		'## Forum_SQL
		strForumSql = "SELECT FORUM_ID "
		strForumSql = strForumSql & " FROM " & strTablePrefix & "FORUM "
	
		Set rsForums = Server.CreateObject("ADODB.Recordset")
		rsForums.open strForumSql, my_Conn
	
		while not 	rsForums.EOF and not rsForums.BOF
			fNum = rsForums("FORUM_ID")
			if not isnumeric(fNum) then fNum = 0
			if chkForumRSSAccess(fNum,submittedUsrID,submittedChkID) then
				strAllowedForums = strAllowedForums & cstr(fNum) & ","
			end if
			rsForums.MoveNext
		wend
		rsForums.close
		set rsForums = nothing
		
		if strAllowedForums <> "" and strAllowedForums <> "," then
			strAllowedForums = left(strAllowedForums,len(strAllowedForums)-1)  ' remove the extra comma
			strAllowedForums = " AND F.FORUM_ID IN (" & strAllowedForums & ")"
		Else
			strAllowedForums = " AND F.F_PRIVATEFORUMS = 0"
		end if
	else
		submittedUsrID = 0
		strAllowedForums = " AND F.F_PRIVATEFORUMS = 0"
	end if
else
	submittedUsrID = 0
	strAllowedForums = " AND F.F_PRIVATEFORUMS = 0"
end if

strSql = "SELECT "
strSql = strSql & " T.T_REPLIES,"
strSql = strSql & " T.T_SUBJECT,"
strSql = strSql & " T.TOPIC_ID,"
strSql = strSql & " T.T_LAST_POST,"
strSql = strSql & " T.T_LAST_POST_AUTHOR,"
strSql = strSql & " T.T_LAST_POST_REPLY_ID,"
strSql = strSql & " T.T_MESSAGE"
strSql = strSql & " FROM " & strTablePrefix & "TOPICS T," & strTablePrefix & "FORUM F "
strSql = strSql & " WHERE T.FORUM_ID = F.FORUM_ID"
'#### strSql = strSql & " AND F.F_PRIVATEFORUMS = 0"  #### replace this with custom list of allowed forums
strSql = strSql & strAllowedForums
If IsNumeric(Request.QueryString("FORUM_ID")) And Request.QueryString("FORUM_ID") > 0 Then
	strSql = strSql & " AND T.FORUM_ID = " & cLng(Request.QueryString("FORUM_ID"))
End If
If IsNumeric(Request.QueryString("CAT_ID")) And Request.QueryString("CAT_ID") > 0 Then
	strSql = strSql & " AND T.CAT_ID = " & cLng(Request.QueryString("CAT_ID"))
End If
strSql = strSql & " AND T.T_STATUS = 1"
strSql = strSql & " ORDER BY T_LAST_POST DESC"

strSql = TopSQL(strSQL, intResults)

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
xml = "<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>" & vbNewLine
xml = xml & "<!-- RSS generation done by Snitz Forums 2000 on " & chkDate(datetostr(strForumTimeAdjust)," ",true) & " -->" & vbNewLine
xml = xml & "<rss version=""2.0"">" & vbNewLine
xml = xml & "<channel>" & vbNewLine
xml = xml & "<language>en-us</language>" & vbNewLine
xml = xml & "<lastBuildDate>" & Date2RFC822(strForumTimeAdjust)& "</lastBuildDate>" & vbNewLine
xml = xml & "<webMaster>" & strSender & "</webMaster>" & vbNewLine
xml = xml & "<ttl>60</ttl>" & vbNewLine
'#### get title 
if Request.QueryString("FORUM_ID") = "" AND Request.QueryString("CAT_ID") = "" then
	strNewTitle = strForumTitle
else
	if Request.QueryString("FORUM_ID") <> "" then
			strTempForum = cLng(request.querystring("FORUM_ID"))
				strsql = "SELECT F_SUBJECT FROM " & strTablePrefix & "FORUM WHERE FORUM_ID=" & strTempForum
				set tforums = my_conn.execute(strsql)
				if tforums.bof or tforums.eof then
					strNewTitle = strForumTitle
					set tforums = nothing
				else	
					strTempForumTitle = chkString(tforums("F_SUBJECT"),"display")
					set tforums = nothing
					strNewTitle = strForumTitle & " - " & strTempForumTitle
				end if
	else
			strTempCat = cLng(request.querystring("CAT_ID"))
				strsql = "SELECT CAT_NAME FROM " & strTablePrefix & "CATEGORY WHERE CAT_ID=" & strTempCat
				set tCat = my_conn.execute(strsql)
				if tCat.bof or tCat.eof then
					strNewTitle = strForumTitle 
					set tCat = nothing
				else	
					strTempForumTitle = chkString(tCat("CAT_NAME"),"display")
					set tCat = nothing
					strNewTitle = strForumTitle & " - " & strTempForumTitle
				end if
	end if
end if

xml = xml & "<title>" & strNewTitle & "</title>" & vbNewLine
xml = xml & "<link>" & strForumURL & "</link>" & vbNewLine
xml = xml & "<description>" & strForumTitle & " " & strTitleOwner & "</description>" & vbNewLine
'xml = xml & "<author>Eastover Fire Department</author>"
xml = xml & "<image>" & vbNewLine
xml = xml & "<link>" & strForumURL & "</link>" & vbNewLine
xml = xml & "<url>" & strForumURL & "xmlfeed.gif</url>" & vbNewLine
xml = xml & "<title>" & strForumTitle & " RSS Feed</title>" & vbNewLine
xml = xml & "<width>144</width>" & vbNewLine
xml = xml & "<height>47</height>" & vbNewLine
xml = xml & "</image>" & vbNewLine

if recActiveTopicsCount <> "" then
	fT_REPLIES = 0
	fT_SUBJECT = 1
	fTOPIC_ID = 2
	fT_LAST_POST = 3
	fT_LAST_POST_AUTHOR = 4
	fT_LAST_POST_REPLY_ID = 5
	fT_MESSAGE = 6

	for RowCount = 0 to recActiveTopicsCount
		Topic_Replies = allActiveTopics(fT_REPLIES,RowCount)
		Topic_Subject = chkstring(replace(allActiveTopics(fT_SUBJECT,RowCount),"&","&amp;"),"display")
		Topic_ID = allActiveTopics(fTOPIC_ID,RowCount)
		Topic_Last_Post = allActiveTopics(fT_LAST_POST,RowCount)
		Topic_Last_Post_Author = getMemberName(allActiveTopics(fT_LAST_POST_AUTHOR,RowCount))
		Topic_Last_Post_Reply_ID = allActiveTopics(fT_LAST_POST_REPLY_ID,RowCount)
		
		if Topic_Replies = 1 then
		    Body = "There is " & Topic_Replies & " reply, posted on " & chkDate(Topic_Last_Post," at",true) & " by " & Topic_Last_Post_Author
		ElseIf Topic_Replies > 0 then
		    Body = "There are " & Topic_Replies & " replies, with the last one on " & chkDate(Topic_Last_Post," at",true) & " by " & Topic_Last_Post_Author
		else
		    'Body = allActiveTopics(fT_MESSAGE, RowCount)
		    Body = "This is a new topic posted on " & chkDate(Topic_Last_Post," at",true) & " by " & Topic_Last_Post_Author & ":<br /><br />" & allActiveTopics(fT_MESSAGE, RowCount)
		end if
		
		Body = MakeCData(Body)
		

		xml = xml & "<item>"
		xml = xml & "<title>" & Topic_Subject & "</title>"
		xml = xml & "<author>" & Topic_Last_Post_Author & "@nospam.com</author>"
		xml = xml & "<link>" & strForumURL & DoLastPostLink & "</link>"
		xml = xml & "<category>" & Forum_Subject & "</category>" & vbNewLine 
		xml = xml & "<pubDate>"& Date2RFC822(StrToDate(Topic_Last_Post)) &"</pubDate>"
		xml = xml & "<guid>" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & "</guid>" & vbNewLine
		xml = xml & "<description>" & Body & "</description>"
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

Function Date2RFC822(Date2Convert)
		'convert the date to the RFC-822 format
		'first declare the variables used:
		dim rfc822timezone,rfc822daydate,rfc822dayno,rfc822day,rfc822monthno,rfc822month,rfc822year,rfc822hour,rfc822minute,rfc822seconds,rfc822time,pubdate
		'first we get the input date
		'Date2Convert = chkDate(Topic_Last_Post,"",true)
		' define your timezone offset below. Examples : "+0100" for GMT+1, "EST", "GMT"
		rfc822timezone = " -0500"
		'get the date (day)
		rfc822daydate = Day(Date2Convert)
		if len(rfc822daydate) = 1 then rfc822daydate = "0" & rfc822daydate
		'get the number of the day of the week, assuming that monday is the first day of the week.
		rfc822dayno = Weekday(Date2Convert, 2)
		' now make sure that this day is translated into the correct english abbreviation:
			select case rfc822dayno
					case 1
						rfc822day = "Mon"
					case 2 
						rfc822day = "Tue"
					case 3 
						rfc822day = "Wed"
					case 4 
						rfc822day = "Thu"
					case 5 
						rfc822day = "Fri"
					case 6 
						rfc822day = "Sat"
					case 7 
						rfc822day = "Sun"
			end select
		rfc822monthno = Month(Date2Convert)
		' now make sure that this month is translated into the correct english abbreviation:
			select case rfc822monthno
				case 1
					rfc822month = "Jan"
				case 2
					rfc822month = "Feb"
				case 3 
					rfc822month = "Mar"
				case 4 
					rfc822month = "Apr"
				case 5
					rfc822month = "May"
				case 6
					rfc822month = "Jun"
				case 7
					rfc822month = "Jul"
				case 8
					rfc822month = "Aug"
				case 9
					rfc822month = "Sep"
				case 10
					rfc822month = "Oct"
				case 11
					rfc822month = "Nov"
				case 12
					rfc822month = "Dec"
			end select
		rfc822year = Year(Date2Convert)
		rfc822hour = Hour(Date2Convert) & ":"
			if len(rfc822hour) = 2 then
			rfc822hour = "0" & rfc822hour
			end if
		rfc822minute = Minute(Date2Convert) & ":"
			if len(rfc822minute) = 2 then
			rfc822minute = "0" & rfc822minute
			end if
		rfc822seconds = second(Date2Convert)
			if len(rfc822seconds) = 1 then
			rfc822seconds = "0" & rfc822seconds
			end if
		rfc822time = rfc822hour & rfc822minute & rfc822seconds
		'now pu the whole thing together in the RFC822 format
		'Example Tue, 21 Dec 2004 22:41:31 +0100
		'Example : DDD, dd MMM yyyy, hh:mm:ss timezone
		Date2RFC822 = rfc822day & ", " & rfc822daydate & " " & rfc822month & " " & rfc822year & " " & rfc822time & rfc822timezone
		'done
end Function

Function DoLastPostLink()
	if Topic_Replies < 1 or Topic_Last_Post_Reply_ID = 0 then
		DoLastPostLink = "topic.asp?TOPIC_ID=" & Topic_ID
	elseif Topic_Last_Post_Reply_ID <> 0 then
		PageLink = "whichpage=-1&amp;"
		AnchorLink = "&amp;REPLY_ID="
		DoLastPostLink = "topic.asp?" & PageLink & "TOPIC_ID=" & Topic_ID & AnchorLink & Topic_Last_Post_Reply_ID
	else
		DoLastPostLink = "topic.asp?TOPIC_ID=" & Topic_ID
	end if
end function

Function GetReplyBody()
    strSqlReq = "SELECT R_MESSAGE FROM " & _
             strTablePrefix & "REPLY WHERE " & _
             "  REPLY_ID=" & Topic_Last_Post_Reply_ID
             
    set nrs = Server.CreateObject("ADODB.Recordset")
    nrs.open strSqlReq, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    if not nrs.EOF then
	    reply = nrs.GetRows(adGetRowsRest)
    end if
    
    nrs.close
    set nrs = nothing
    
    GetReplyBody = reply(0,0)
end function

Function MakeCData( foo ) 
    MakeCData = "<![CDATA[" & formatStr(foo) & "]]>"
end function

%>