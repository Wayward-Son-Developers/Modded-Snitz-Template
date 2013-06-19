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
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
intStep = Request.QueryString("Step")
if intStep = "" or IsNull(intStep) then
	intStep = 1
else
	intStep = cLng(intStep)
end if
if intStep < 5 then
	Select Case LCase(request.querystring("comeback"))
		Case "etc"
			Response.write("<meta http-equiv=""Refresh"" content=""1; URL=admin_count.asp?Step=" & intStep + 1 & "&comeback=etc"">")
		Case Else
			Response.write "<meta http-equiv=""Refresh"" content=""1; URL=admin_count.asp?Step=" & intStep + 1 & """>"
	End Select
else
	Select Case LCase(request.querystring("comeback"))
		Case "etc"
			Response.write("<meta http-equiv=""Refresh"" content=""1; URL=admin_etc.asp?action=deletememtopicsdone"">")
		Case Else
			Response.write "<meta http-equiv=""Refresh"" content=""60; URL=admin_home.asp"">"
	End Select
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Update&nbsp;Forum&nbsp;Counts</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"<table class=""admin"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td colspan=""2"">Updating Counts Step " & intStep & " of 5</td>" & vbNewLine & _
				"</tr>" & vbNewLine

Server.ScriptTimeout = 6000

Select Case intStep
	Case 1
		Response.Write "<tr>" & vbNewline
		Response.Write "<td colspan=""2"">Topics" & vbNewline
	
		'## Forum_SQL - Get contents of the Forum table related to counting
		strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
		if rs.EOF then
			recForumCount = ""
		else
			allForumData = rs.GetRows(adGetRowsRest)
			recForumCount = UBound(allForumData,2)
		end if
	
		rs.close
		set rs = nothing
	
		if recForumCount <> "" then
			fFORUM_ID = 0
			fF_TOPICS = 1
			i = 0 
	
			for iForum = 0 to recForumCount
				ForumID = allForumData(fFORUM_ID,iForum)
				ForumTopics = allForumData(fF_TOPICS,iForum)
	
				i = i + 1
	
				'## Forum_SQL - count total number of topics in each forum in Topics table
				strSql = "SELECT count(FORUM_ID) AS cnt "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE FORUM_ID = " & ForumID
				strSql = strSql & " AND T_STATUS <= 1 "
	
				set rs1 = my_Conn.Execute(strSql)
	
				if rs1.EOF or rs1.BOF then
					intF_TOPICS = 0
				else
					intF_TOPICS = rs1("cnt")
				end if
	
				set rs1 = nothing
	
				'## Forum_SQL - count total number of topics in each forum in A_Topics table
				strSql = "SELECT count(FORUM_ID) AS cnt "
				strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
				strSql = strSql & " WHERE FORUM_ID = " & ForumID
				strSql = strSql & " AND T_STATUS <= 1 "
	
				set rs1 = my_Conn.Execute(strSql)
	
				if rs1.EOF or rs1.BOF then
					intF_A_TOPICS = 0
				else
					intF_A_TOPICS = rs1("cnt")
				end if
	
				set rs1 = nothing
	
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET F_TOPICS = " & intF_TOPICS
				strSql = strSql & " , F_A_TOPICS = " & intF_A_TOPICS
				strSql = strSql & " WHERE FORUM_ID = " & ForumID
	
				my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
	
				Response.Write "."
				if i = 80 then 
					Response.Write "<br />" & vbNewline
					i = 0
				end if
			next
		end if
	
		Response.Write "</td>" & vbNewline
		Response.Write "</tr>" & vbNewline
		
	Case 2
		Response.Write "<tr>" & vbNewline
		Response.Write "<td colspan=""2"">Topic Replies" & vbNewline
	
		'## Forum_SQL
		strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS"
		strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.T_STATUS <= 1"
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
		if rs.EOF then
			recTopicCount = ""
		else
			allTopicData = rs.GetRows(adGetRowsRest)
			recTopicCount = UBound(allTopicData,2)
		end if
	
		rs.close
		set rs = nothing
	
		if recTopicCount <> "" then
			fTOPIC_ID = 0
			fT_REPLIES = 1
			i = 0 
	
			for iTopic = 0 to recTopicCount
				TopicID = allTopicData(fTOPIC_ID,iTopic)
				TopicReplies = allTopicData(fT_REPLIES,iTopic)
	
				i = i + 1
	
				'## Forum_SQL - count total number of replies in Topics table
				strSql = "SELECT count(REPLY_ID) AS cnt "
				strSql = strSql & " FROM " & strTablePrefix & "REPLY "
				strSql = strSql & " WHERE TOPIC_ID = " & TopicID
				strSql = strSql & " AND R_STATUS <= 1 "
	
				set rs = Server.CreateObject("ADODB.Recordset")
				rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
				if rs.EOF then
					recReplyCntCount = ""
				else
					allReplyCntData = rs.GetRows(adGetRowsRest)
					recReplyCntCount = UBound(allReplyCntData,2)
				end if
	
				rs.close
				set rs = nothing
	
				if recReplyCntCount <> "" then
					fReplyCnt = 0
	
					for iCnt = 0 to recReplyCntCount
						ReplyCnt = allReplyCntData(fReplyCnt,iCnt)
	
						intT_REPLIES = ReplyCnt
	
						'## Forum_SQL - Get last_post and last_post_author for Topic
						strSql = "SELECT R_DATE, R_AUTHOR "
						strSql = strSql & " FROM " & strTablePrefix & "REPLY "
						strSql = strSql & " WHERE TOPIC_ID = " & TopicID & " "
						strSql = strSql & " AND R_STATUS <= 1"
						strSql = strSql & " ORDER BY R_DATE DESC"
	
						set rs2 = my_Conn.Execute (strSql)
	
						if not(rs2.eof or rs2.bof) then
							rs2.movefirst
							strLast_Post = rs2("R_DATE")
							strLast_Post_Author = rs2("R_AUTHOR")
						else
							strLast_Post = ""
							strLast_Post_Author = ""
						end if
	
						set rs2 = nothing
					next
							else
					intT_REPLIES = 0
	
					set rs2 = Server.CreateObject("ADODB.Recordset")
	
					'## Forum_SQL - Get post_date and author from Topic
					strSql = "SELECT T_AUTHOR, T_DATE "
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE TOPIC_ID = " & TopicID & " "
					strSql = strSql & " AND T_STATUS <= 1"
	
					set rs2 = my_Conn.Execute(strSql)
	
					if not(rs2.eof or rs2.bof) then
						strLast_Post = rs2("T_DATE")
						strLast_Post_Author = rs2("T_AUTHOR")
					else
						strLast_Post = ""
						strLast_Post_Author = ""
					end if
	
					set rs2 = nothing
	
				end if
	
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
				if strLast_Post <> "" then 
					strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
					if strLast_Post_Author <> "" then 
						strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author 
					end if
				end if
				strSql = strSql & " WHERE TOPIC_ID = " & TopicID
	
				my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
	
				Response.Write "."
				if i = 80 then 
					Response.Write "<br />" & vbNewline
					i = 0
				end if
			next
		end if
	
		Response.Write "</td>" & vbNewline
		Response.Write "</tr>" & vbNewline
		
	Case 3
		Response.Write "<tr>" & vbNewline
		Response.Write "<td colspan=""2"">UnModerated Topic Replies" & vbNewline
	
		'## Forum_SQL
		strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS"
		strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.T_STATUS <= 1"
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
		if rs.EOF then
			recTopicCount = ""
		else
			allTopicData = rs.GetRows(adGetRowsRest)
			recTopicCount = UBound(allTopicData,2)
		end if
	
		rs.close
		set rs = nothing
	
		if recTopicCount <> "" then
			fTOPIC_ID = 0
			i = 0 
	
			for iTopic = 0 to recTopicCount
				TopicID = allTopicData(fTOPIC_ID,iTopic)
	
				i = i + 1
	
				'## Forum_SQL - count total number of unmoderated replies in Topics table
				strSql = "SELECT count(REPLY_ID) AS cnt "
				strSql = strSql & " FROM " & strTablePrefix & "REPLY "
				strSql = strSql & " WHERE TOPIC_ID = " & TopicID
				strSql = strSql & " AND R_STATUS = 2 OR R_STATUS = 3 "
	
				set rs = Server.CreateObject("ADODB.Recordset")
				rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
				if rs.EOF then
					recReplyCntCount = ""
				else
					allReplyCntData = rs.GetRows(adGetRowsRest)
					recReplyCntCount = UBound(allReplyCntData,2)
				end if
	
				rs.close
				set rs = nothing
	
				if recReplyCntCount <> "" then
					fReplyCnt = 0
					for iCnt = 0 to recReplyCntCount
						intT_UREPLIES = allReplyCntData(fReplyCnt,iCnt)
					next
							else
					intT_UREPLIES = 0
				end if
	
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_UREPLIES = " & intT_UREPLIES
				strSql = strSql & " WHERE TOPIC_ID = " & TopicID
	
				my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
	
				Response.Write "."
				if i = 80 then 
					Response.Write "<br />" & vbNewline
					i = 0
				end if
			next
		end if
	
		Response.Write "</td>" & vbNewline
		Response.Write "</tr>" & vbNewline
		
	Case 4
		Response.Write "<tr>" & vbNewline
		Response.Write "<td colspan=""2"">Forum Replies" & vbNewline
	
		'## Forum_SQL - Get values from Forum table needed to count replies
		strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
		if rs.EOF then
			recForumCount = ""
		else
			allForumData = rs.GetRows(adGetRowsRest)
			recForumCount = UBound(allForumData,2)
		end if
	
		rs.close
		set rs = nothing
	
		if recForumCount <> "" then
			fFORUM_ID = 0
			fF_COUNT = 1
			i = 0
	
			for iForum = 0 to recForumCount
				ForumID = allForumData(fFORUM_ID,iForum)
				ForumCount = allForumData(fF_COUNT,iForum)
	
				i = i + 1
	
				'## Forum_SQL - Count total number of Replies
				strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & ForumID
				strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"
	
				set rs1 = my_Conn.Execute(strSql)
	
				if rs1.EOF or rs1.BOF then
					intF_COUNT = 0
					intF_TOPICS = 0
				else
					if IsNull(rs1("SumOfT_REPLIES")) then
						intF_COUNT = rs1("cnt")
					else
						intF_COUNT = CLng(rs1("cnt")) + CLng(rs1("SumOfT_REPLIES"))
					end if
					intF_TOPICS = rs1("cnt") 
				end if
				if IsNull(intF_COUNT) then intF_COUNT = 0 
				if IsNull(intF_TOPICS) then intF_TOPICS = 0 
	
				set rs1 = nothing
	
				'## Forum_SQL - Count total number of Archived Replies
				strSql = "SELECT Sum(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS cnt "
				strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
				strSql = strSql & " WHERE " & strTablePrefix & "A_TOPICS.FORUM_ID = " & ForumID
				strSql = strSql & " AND " & strTablePrefix & "A_TOPICS.T_STATUS <= 1"
	
				set rs1 = my_Conn.Execute(strSql)
	
				if rs1.EOF or rs1.BOF then
					intF_A_COUNT = 0
					intF_A_TOPICS = 0
				else
					if IsNull(rs1("SumOfT_REPLIES")) then
						intF_A_COUNT = rs1("cnt")
					else
						intF_A_COUNT = CLng(rs1("cnt")) + CLng(rs1("SumOfT_REPLIES"))
					end if
					intF_A_TOPICS = rs1("cnt") 
				end if
				if IsNull(intF_A_COUNT) then intF_A_COUNT = 0 
				if IsNull(intF_A_TOPICS) then intF_A_TOPICS = 0 
	
				set rs1 = nothing
	
				'## Forum_SQL - Get last_post and last_post_author for Forum
				strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE FORUM_ID = " & ForumID & " "
				strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"
				strSql = strSql & " ORDER BY T_LAST_POST DESC"
	
				set rs2 = my_Conn.Execute (strSql)
	
				if not (rs2.eof or rs2.bof) then
					strLast_Post = rs2("T_LAST_POST")
					strLast_Post_Author = rs2("T_LAST_POST_AUTHOR")
				else
					strLast_Post = ""
					strLast_Post_Author = ""
				end if
	
				set rs2 = nothing
	
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET F_COUNT = " & intF_COUNT
				strSql = strSql & ",  F_TOPICS = " & intF_TOPICS
				strSql = strSql & ",  F_A_COUNT = " & intF_A_COUNT
				strSql = strSql & ",  F_A_TOPICS = " & intF_A_TOPICS
				if strLast_Post <> "" then 
					strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
					if strLast_Post_Author <> "" then 
						strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
					end if
				end if
				strSql = strSql & " WHERE FORUM_ID = " & ForumID
	
				my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
	
				Response.Write "."
				if i = 80 then 
					Response.Write "<br />" & vbNewline
					i = 0
				end if	
			next
		end if
		Response.Write "</td>" & vbNewline
		Response.Write "</tr>" & vbNewline
		
	Case 5
		Response.Write "<tr>" & vbNewline
		Response.Write "<td class=""formlabel"">Totals:</td>" & vbNewline
		Response.Write "<td class=""formvalue"">"
	
		'## Forum_SQL - Total of Topics
		strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
		strSql = strSql & " AS SumOfF_TOPICS "
		strSql = strSql & ", Sum(" & strTablePrefix & "FORUM.F_A_TOPICS) "
		strSql = strSql & " AS SumOfF_A_TOPICS "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	
		set rs = my_Conn.Execute(strSql)
	
		if rs("SumOfF_TOPICS") <> "" then
			Response.Write "Total Topics: " & rs("SumOfF_TOPICS") & "<br />" & vbNewline
			intSumOfF_TOPICS = rs("SumOfF_TOPICS")
		else
			Response.Write "Total Topics: 0<br />" & vbNewLine
			intSumOfF_TOPICS = 0
		end if
		if rs("SumOfF_A_TOPICS") <> "" then
			Response.Write "Archived Topics: " & rs("SumOfF_A_TOPICS") & "<br />" & vbNewline
			intSumOfF_A_TOPICS = rs("SumOfF_A_TOPICS")
		else
			Response.Write "Archived Topics: 0<br />" & vbNewLine
			intSumOfF_A_TOPICS = 0
		end if
	
		'## Forum_SQL - Write total Topics to Totals table
		strSql = "UPDATE " & strTablePrefix & "TOTALS "
		strSql = strSql & " SET T_COUNT = " & intSumOfF_TOPICS
		strSql = strSql & " , T_A_COUNT = " & intSumOfF_A_TOPICS
	
		set rs = nothing
	
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	
		'## Forum_SQL - Total all the replies for each topic
		strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
		strSql = strSql & " AS SumOfF_COUNT "
		strSql = strSql & ", Sum(" & strTablePrefix & "FORUM.F_A_COUNT) "
		strSql = strSql & " AS SumOfF_A_COUNT "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	
		set rs = my_Conn.Execute (strSql)
	
		if rs("SumOfF_COUNT") <> "" then
			Response.Write "Total Posts: " & RS("SumOfF_COUNT") & "<br />" & vbNewline
			intSumOfF_COUNT = rs("SumOfF_COUNT")
		else
			Response.Write "Total Posts: 0<br />" & vbNewline
			intSumOfF_COUNT = 0
		end if
		if rs("SumOfF_A_COUNT") <> "" then
			Response.Write "Total Archived Posts: " & RS("SumOfF_A_COUNT") & "<br />" & vbNewline
			intSumOfF_A_COUNT = rs("SumOfF_A_COUNT")
		else
			Response.Write "Total Posts: 0<br />" & vbNewline
			intSumOfF_A_COUNT = 0
		end if
	
		'## Forum_SQL - Write total replies to the Totals table
		strSql = "UPDATE " & strTablePrefix & "TOTALS "
		strSql = strSql & " SET P_COUNT = " & intSumOfF_COUNT
		strSql = strSql & " , P_A_COUNT = " & intSumOfF_A_COUNT
	
		set rs = nothing
	
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	
		'## Forum_SQL - Total number of users
		strSql = "SELECT Count(MEMBER_ID) "
		strSql = strSql & " AS CountOf "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	
		set rs = my_Conn.Execute(strSql)
	
		Response.Write "Registered Users: " & rs("Countof") & "<br />" & vbNewline
	
		'## Forum_SQL - Write total number of users to Totals table
		strSql = " UPDATE " & strTablePrefix & "TOTALS "
		strSql = strSql & " SET U_COUNT = " & cLng(RS("Countof"))
	
		set rs = nothing
	
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	
		Response.Write "</td>" & vbNewline
		Response.Write "</tr>" & vbNewline
		Response.Write "<tr class=""header"">" & vbNewline
		Response.Write "<td colspan=""2"">Count Update Complete</td>" & vbNewline
		Response.Write "</tr>" & vbNewline
		Response.Write "<tr>" & vbNewline
		Response.Write "<td class=""footernav"" colspan=""2""><a href=""admin_home.asp"">Back to Admin Home</a></td>" & vbNewline
		Response.Write "</tr>" & vbNewline
End Select


response.write "</table>"

WriteFooter
Response.End
%>
