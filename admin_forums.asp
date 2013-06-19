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
	Response.Redirect "admin_login.asp?target=" & server.urlencode(scriptname(ubound(scriptname)) & "?" & request.querystring)
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
		"<tr>" & vbNewLine & _
		"<td class=""secondnav"">" & vbNewLine & _
		getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
		getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br />" & vbNewLine & _
		getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;<a href=""admin_forums.asp"">Forum Archive &amp; Cleanup</a></td>" & vbNewLine & _
		"</tr>" & vbNewLine & _
		"</table>" & vbNewLine

strWhatToDo = request.querystring("action")
Select Case strWhatToDo
	Case "deleteftopics"
		strForumIDN = request.querystring("id")
		strForumIDN = Server.URLEncode(strForumIDN)
		if strForumIDN = "" then
			strsql = "SELECT CAT_ID, FORUM_ID, F_L_DELETE, F_SUBJECT,F_DELETE_SCHED FROM " & strTablePrefix & "FORUM ORDER BY CAT_ID, F_SUBJECT DESC"
			
			Response.Write	"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td>Administrative Forum Delete Functions</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr class=""section"">" & vbNewLine & _
							"<td>Delete Topics</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td>" & vbNewLine
			set drs = my_conn.execute(strsql)    
			thisCat = 0
			if drs.eof then
				Response.write	"No Forums Found" & vbNewLine
			else
				Response.write	"<ul>" & vbNewLine & _
								"<li><a href=""admin_forums.asp?action=deleteftopics&id=-1"">All Forums</a></li>" & vbNewLine & _
								"<li><a href=""javascript:document.delTopic.submit()"">Selected Forums</a></li>" & vbNewLine & _
								"<li><form name=""delTopic"" action=""admin_forums.asp"">" & vbNewLine & _
								"<input type=""hidden"" value=""deleteftopics"" name= ""action"" >" & vbNewLine & _
								"<ul>" & vbNewLine
				do until drs.eof
					lastDeleted = drs("F_L_DELETE")
					schedDays = drs("F_DELETE_SCHED")
					if (IsNull(lastDeleted)) or (lastDeleted = "") then 
						delete_date = "N/A"
						overdue = 0
					else 
						needDelete = (DateAdd("d",schedDays+7,strToDate(lastDeleted)))
						if (strForumTimeAdjust > needDelete) and (schedDays > 0) then
							overdue = true
							delete_date = "<span class=""hlf"">Deletion Overdue</span>"
						else
							overdue = false
							delete_date = StrToDate(lastDeleted)
						end if
					end if
					Response.Write	"<li><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & """"
					if overdue then Response.Write(" checked")
					Response.Write	">&nbsp;<a href=""admin_forums.asp?action=deleteftopics&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a>&nbsp;" & _
									"(Last delete date: " & delete_date & ")</li>" & vbNewLine
					thisCat = drs("CAT_ID")
					drs.movenext
				loop
				Response.Write	"</ul></form></li></ul>" & vbNewLine
			end if
			Response.Write	"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
			set drs = nothing
		else
			Select Case LCase(request.querystring("confirm"))
				Case "true"
					Call subdeletestuff(strForumIDN)
					Call OkMessage("All Topics in selected Forum/s have been Deleted.","admin_forums.asp","Back to Forum Archive Functions")
				Case "false"
					Call OkMessage("Topics in selected Forum(s) have NOT been deleted.","admin_forums.asp","Back to Forum Archive Functions")
				Case Else
					Response.Write	"<div class=""warning"" style=""width:50%;"">Are you sure you want to delete <b>ALL</b> topics"
					if Request.QueryString("id") = "-1" then Response.Write(" in <b>ALL</b> forums? ") else Response.Write(" in the selected forums? ")
					Response.Write	"This is <B><STRONG>NOT</STRONG></B> reversible.<br />" & vbNewLine & _
									"<a href=""admin_forums.asp?action=deleteftopics&id=" & strForumIDN & "&confirm=true"">Yes</a> | <a href=""admin_forums.asp?action=deleteftopics&id=" & strForumIDN & "&confirm=false"">No</a></div>" & vbNewLine
			End Select
		end if
	Case "archive"
		strForumIDN = request("id")
		strForumIDN = Server.URLEncode(strForumIDN)
		if strForumIDN = "" then
			Response.Write	"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td>Administrative Forum Archive Functions</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr class=""section"">" & vbNewLine & _
							"<td>Archive All Topics</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td>" & vbNewLine
			
			strsql = "Select CAT_ID, FORUM_ID, F_L_ARCHIVE, F_SUBJECT,F_ARCHIVE_SCHED from " & strTablePrefix & "FORUM WHERE F_TYPE = 0 ORDER BY CAT_ID, F_SUBJECT DESC"
			
			set drs = my_conn.execute(strsql)    
			thisCat = 0
			if drs.eof then
				Response.write	"No Forums Found" & vbNewLine
			else
				Response.Write	"<ul>" & vbNewLine & _
								"<li><a href=""admin_forums.asp?action=archive&id=-1"">All Forums</a></li>" & vbNewLine & _
								"<li><a href=""javascript:document.arcTopic.submit()"">Selected Forums</a></li>" & vbNewLine & _
								"<li><ul>" & vbNewLine & _
								"<form name=""arcTopic"" action=""admin_forums.asp"">" & vbNewLine & _
								"<input type=""hidden"" value=""archive"" name=""action"">" & vbNewLine
				do until drs.eof
					lastArchived = drs("F_L_ARCHIVE")
					schedDays = drs("F_ARCHIVE_SCHED")
					if (IsNull(lastArchived)) or (lastArchived = "") then 
						archive_date = "Not archived" 
						overdue = 0
					else 
						needArchive = (DateAdd("d",schedDays+7,strToDate(lastArchived)))
						if (strForumTimeAdjust > needArchive) and (schedDays > 0) then
							overdue = true
							archive_date = "<span class=""hlf"">Archiving Overdue</span>"
						else
							overdue = false
							archive_date = StrToDate(lastArchived)
						end if
					end if
					Response.Write	"<li><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & """"
					if overdue then Response.Write(" checked")
					Response.Write	""">&nbsp;<a href=""admin_forums.asp?action=archive&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a>&nbsp;" & vbNewLine & _
									"(Last archive date: " & archive_date & ")</li>" & vbNewLine
					thisCat = drs("Cat_ID")
					drs.movenext
				loop
				Response.Write	"</ul></form></li></ul>" & vbNewLine
			end if
			set drs = nothing
			Response.Write	"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		elseif strForumIDN <> "" then
			Select Case LCase(request.querystring("confirm"))
				Case "no"
					Response.Write	"<div class=""warning"" style=""width:50%;"">Are you sure you want to archive these topics?<br />" & vbNewline & _
									"<a href=""admin_forums.asp?action=archive&id=" & strForumIDN & "&confirm=yes&date=" & request.form("archiveolderthan") & """>Yes</a> | <a href=""admin_forums.asp?action=archive&id=" & strForumIDN & "&confirm=cancel"">No</a></div>" & vbNewLine
				Case "yes"
					If chkDateFormat(request.querystring("date")) Then Call subarchivestuff(request.querystring("date"))
				Case "cancel"
					Call OkMessage("Archiving Cancelled","admin_forums.asp","Back to Forum Archive Functions")
				Case Else
					Response.Write	"<table class=""admin"">" & vbNewLine & _
									"<tr class=""header"">" & vbNewLine & _
									"<td>Administrative Forum Archive Functions</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr class=""section"">" & vbNewLine & _
									"<td>Archive All Topics</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td>" & vbNewLine & _
									"<form method=""post"" action=""admin_forums.asp?action=archive&id=" & strForumIDN & "&confirm=no"">" & vbNewLine & _
									"Archive Topics which are older than:<br />" & vbNewLine & _
									"<select name=""archiveolderthan"" size=""1"">" & vbNewLine
					for counter = 1 to 6
						Response.Write	"<option value=""" & DateToStr(DateAdd("m", -counter, now())) & """>" & counter & " Month"
						if counter > 1 then response.write("s")
						Response.Write	"</option>" & vbNewLine
					next
					Response.Write	"<option value=""" & DateToStr(DateAdd("m", -12, now())) & """>One Year</option>" & vbNewLine & _
									"</select>&nbsp;<button type=""submit"">Archive</button>" & vbNewLine & _
									"</form>" & vbNewLine & _
									"</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"</table>" & vbNewLine
			End Select
		end if
	Case "deletearchive"
		strForumIDN = request.querystring("id")
		strForumIDN = Server.URLEncode(strForumIDN)
		if strForumIDN = "" then
   			strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, "
		    strSql = strSql & strTablePrefix & "FORUM.FORUM_ID, "
		    strSql = strSql & strTablePrefix & "FORUM.F_L_DELETE, "   
		    strSql = strSql & strTablePrefix & "FORUM.F_DELETE_SCHED, "
		    strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT "
		    strSql = strSql & " FROM " & strTablePrefix & "FORUM, " & strArchiveTablePrefix & "TOPICS " 
		    strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strArchiveTablePrefix & "TOPICS.FORUM_ID "   
		    strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.CAT_ID DESC, " & strTablePrefix & "FORUM.F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)    
			thisCat = 0
			thisForum = 0
			
			
			Response.Write	"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td>Administrative Forum Archive Functions</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr class=""section"">" & vbNewLine & _
							"<td>Delete Archived Topics</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td>" & vbNewLine
			
			if drs.eof then
				Response.write	"No Forums Found!" & vbNewLine
			else
				Response.Write	"<ul>" & vbNewLine & _
								"<li><a href=""admin_forums.asp?action=deletearchive&id=-1"">All Forums</a></li>" & vbNewLine & _
								"<li><a href=""javascript:document.delTopic.submit()"">Selected Forums</a></li>" & vbNewLine & _
								"<li><ul>" & vbNewLine & _
								"<form name=""delTopic"" action=""admin_forums.asp"">" & vbNewLine & _
								"<input type=""hidden"" value=""deletearchive"" name= ""action"">" & vbNewLine
				do until drs.eof
					if thisForum <> drs("FORUM_ID") then
						thisForum = drs("FORUM_ID")
						lastDeleted = drs("F_L_DELETE")
						schedDays = drs("F_DELETE_SCHED")
						
						if (IsNull(lastDeleted)) or (lastDeleted = "") then 
							delete_date = "N/A" 
							overdue = 0
						else 
							needDelete = (DateAdd("d",schedDays+7,strToDate(lastDeleted)))
							if (strForumTimeAdjust > needDelete) and (schedDays > 0) then
								overdue = true
								delete_date = "<span class=""hlf"">Deletion Overdue</span>"
							else
								overdue = false
								delete_date = StrToDate(lastDeleted)
							end if
						end if

						if thisCat <> drs("CAT_ID") then
							thisCat = drs("CAT_ID")
						end if
						Response.Write	"<li><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & ""
						if overdue then Response.Write(" checked")
						Response.Write	""">&nbsp;<a href=""admin_forums.asp?action=deletearchive&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a>&nbsp;" & vbNewLine & _
										"(Last delete date: " & delete_date & ")</li>" & vbNewLine
					end if
					drs.movenext
				loop
				Response.Write	"</ul></form></li></ul>" & vbNewLine
			end if
			set drs = nothing
			Response.Write	"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		elseif request.querystring("id") <> "" then
			Select Case LCase(request.querystring("confirm"))
				Case "no"
					Response.Write	"<div class=""warning"" style=""width:50%;"">Are you sure you want to delete these topics from the archive?<br />" & vbNewline & _
									"<a href=""admin_forums.asp?action=deletearchive&id=" & strForumIDN & "&confirm=yes&date=" & request.form("archiveolderthan") & """>Yes</a> | <a href=""admin_forums.asp?action=delete&confirm=false&id=" & strForumIDN & """>No</a></div>" & vbNewLine
				Case "yes"
					If chkDateFormat(request.querystring("date")) Then
						Call subdeletearchivetopics(strForumIDN, request.querystring("date"))
						Call OkMessage("Topics older than " & StrToDate(request.querystring("date")) & " have been deleted from the selected archive forum.","admin_forums.asp","Back to Forum Archive Functions")
					End If
				Case Else
					Response.Write 	"Select how many months old the Topics should be that you wish to delete:" & vbNewLine & _
									"<form method=""post"" action=""admin_forums.asp?action=deletearchive&id=" & strForumIDN & "&confirm=no"">" & vbNewLine & _
									"<center><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>Delete archived Topics which are older than:</font><br />" & vbNewLine & _
									"<select name=""archiveolderthan"" size=""1"">" & vbNewLine
					for counter = 1 to 6
						Response.Write	"<option value=""" & DateToStr(DateAdd("m", -counter, now())) & """>" & counter & " Month"
						if counter > 1 then Response.Write("s")
						Response.Write	"</option>" & vbNewLine
					next
					Response.Write	"<option value=""" & DateToStr(DateAdd("m", -12, now())) & """>One Year</option>" & vbNewLine & _
									"</select>" & vbNewLine & _
									"&nbsp;&nbsp;" & vbNewLine & _
									"<input type=""submit"" value=""Delete""></center>" & vbNewLine & _
									"</form>" & vbNewLine
			End Select
		end if
	Case Else
		Response.Write	"<table class=""admin"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td>Forum Archive & Cleanup</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>Archive Options:<ul>" & vbNewLine & _
						"<li><a href=""admin_forums_schedule.asp"">Configure Archive Reminder</a></li>" & vbNewLine & _
						"<li><a href=""admin_forums.asp?action=archive"">Archive topics from a forum</a></li>" & vbNewLine & _
						"<li><a href=""admin_forums.asp?action=deletearchive"">Delete selected topics from an archive</a></li>" & vbNewLine & _
						"</ul>" & vbNewLine & _
						"Cleanup Options:<ul>" & vbNewLine & _
						"<li><a href=""admin_forums.asp?action=deleteftopics"">Delete <b>all</b> topics from a forum</a></li>" & vbNewLine & _
						"<li>Delete <b>all</b> topics by a member</li>" & vbNewLine & _
						"</ul>" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine
end Select

Sub subDeleteArchiveTopics(strForum_id, strDateOlderThan)
	Dim fIDSQL
	'#### create FORUM_ID clause
	rqID = request("id")
	'rqID = strForum_id
        on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then 
			fIDSQL = " AND FORUM_ID=" & rqID
		else
			fIDSQL = ""
		end if
		err.clear
	else
		fIDSQL = " AND FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0

	strsql = "DELETE FROM " & strArchiveTablePrefix & "TOPICS WHERE T_LAST_POST < '" & strDateOlderThan & "'" & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	strsql = "DELETE FROM " & strArchiveTablePrefix & "REPLY WHERE R_DATE < '" & strDateOlderThan & "'" & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	Call subdoupdates()
End Sub

Sub subArchiveStuff(fdateolderthan)
	set Server2 = Server
	Server2.ScriptTimeout = 10000
	Dim fIDSQL
	Dim drs,delRep
	
	Set drs = CreateObject("ADODB.Recordset")
	Set delRep = CreateObject("ADODB.Recordset")
	Set drs.ActiveConnection = my_conn
	'#### create FORUM_ID clause
	rqID = request("id")
    	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then 
			fIDSQL = " AND " & strTablePrefix & "TOPICS.FORUM_ID=" & rqID
		else
			fIDSQL = ""
		end if
		err.clear
	else
		fIDSQL = " AND " & strTablePrefix & "TOPICS.FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0
	'#### Get the replies to Archive

	strSql = "SELECT T_DATE, " & strTablePrefix & "REPLY.* FROM " & strTablePrefix & "REPLY LEFT OUTER JOIN " & strTablePrefix & "TOPICS " &_
		 "ON " & strTablePrefix & "REPLY.TOPIC_ID = " & strTablePrefix & "TOPICS.TOPIC_ID " &_
		 " WHERE T_LAST_POST < '" & fdateolderthan & "'" & fIDSQL
	strSQL = strSQL & " AND T_ARCHIVE_FLAG <> 0 "

	drs.Open strsql, my_conn, adOpenStatic, adLockOptimistic, adCmdText

	'#### Archive the Replies
	if drs.eof then
    		response.write("                      <center><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>No Replies were Archived: none found</font></center><br />" & vbNewLine)
	else
        	i = 0
		response.write("                      <font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>")
		do until drs.eof
			if isnull(drs("R_LAST_EDITBY")) then
				intR_LAST_EDITBY = "NULL"
			else
				intR_LAST_EDITBY = drs("R_LAST_EDITBY")
			end if

        		strsqlvalues = "" & drs("CAT_ID") & ", " & drs("FORUM_ID") & ", " & drs("TOPIC_ID") & ", " & drs("REPLY_ID")
		        strsqlvalues = strsqlvalues & ", " & drs("R_AUTHOR") & ", '" & chkstring(drs("R_MESSAGE"),"archive")
	       	        strsqlvalues = strsqlvalues & "', '" & drs("R_DATE") & "', '" & drs("R_IP") & "'"  & ", " & drs("R_STATUS")
			strSqlvalues = strsqlvalues & ", '" & drs("R_LAST_EDIT") & "', " & intR_LAST_EDITBY & ", " & drs("R_SIG") & " "
            
	                strsql = "INSERT INTO " & strArchiveTablePrefix & "REPLY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID, R_AUTHOR, R_MESSAGE, R_DATE, R_IP, R_STATUS, R_LAST_EDIT, R_LAST_EDITBY, R_SIG)"
		        strsql = strsql & " VALUES (" & strsqlvalues & ")"
	
			response.write(".")
			'Response.Write(strSql)
			'Response.End
			my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	           	drs.movenext
			i = i + 1
			if i = 100 then
				response.write("<br />")
				i = 0
			end if
			'#### Delete Original
		Loop
		response.write("</font>" & vbNewLine)
		drs.movefirst
		do while not drs.eof
			strsql = "select * from " & strTablePrefix & "REPLY WHERE REPLY_ID = " & drs("REPLY_ID")
			delrep.Open strsql, my_conn, adOpenStatic, adLockOptimistic, adCmdText
			delrep.delete
			delrep.close
			drs.movenext
		loop

		response.write("                      <center><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>All replies to Topics older than " & strToDate(fdateolderthan) & " were archived</font></center><br />" & vbNewLine)
	end if

	'#### Update FORUM archive date
	strsql = "UPDATE " & strTablePrefix & "FORUM SET F_L_ARCHIVE= '" & fdateolderthan & "'"
	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then 
			strSQL = strSql & " WHERE FORUM_ID=" & rqID
		end if
		err.clear
	else
		strSQL = strSql & " WHERE FORUM_ID IN (" & rqID & ")"
		err.clear
	end if
	on error goto 0
'	strSQL = strSQL & " AND T_ARCHIVE_FLAG <> 0 "

	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords

	'#### Get the TOPICS to Archive
	
	strsql = "SELECT CAT_ID,FORUM_ID,TOPIC_ID,T_SUBJECT,T_AUTHOR,T_REPLIES,T_UREPLIES,T_VIEW_COUNT,T_LAST_POST,T_DATE,T_LAST_POSTER,T_IP,T_LAST_POST_AUTHOR,T_LAST_POST_REPLY_ID,T_LAST_EDIT,T_LAST_EDITBY,T_STICKY,T_SIG,T_MESSAGE FROM " & strTablePrefix & "TOPICS WHERE T_LAST_POST < '" & fdateolderthan & "'" & fIDSQL
	strSQL = strSQL & " AND T_ARCHIVE_FLAG <> 0 "
	set drs = my_conn.execute(strsql)

   
	'#### Archive the Topics
   	if drs.eof then
       		response.write("                      <center><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>No Topics were Archived: none found</font></center><br />" & vbNewLine)
	else
	       	i = 0
       		do until drs.eof
       			strSQL = "SELECT TOPIC_ID FROM " & strArchiveTablePrefix & "TOPICS WHERE TOPIC_ID=" & drs("TOPIC_ID")
			set rsTcheck = my_conn.execute(strSQL)

			if isnull(drs("T_LAST_EDITBY")) then
				intT_LAST_EDITBY = "NULL"
			else
				intT_LAST_EDITBY = drs("T_LAST_EDITBY")
			end if
			if isnull(drs("T_LAST_POST_REPLY_ID")) then
				intT_LAST_POST_REPLY_ID = "NULL"
			else
				intT_LAST_POST_REPLY_ID = drs("T_LAST_POST_REPLY_ID")
			end if
			if isnull(drs("T_UREPLIES")) then
				intT_UREPLIES = "NULL"
				intT_UREPLIEScnt = 0
			else
				intT_UREPLIES = drs("T_UREPLIES")
				intT_UREPLIEScnt = drs("T_UREPLIES")
			end if

			if rsTcheck.eof then
				err.clear

				strsqlvalues = "" & drs("CAT_ID") & ", " & drs("FORUM_ID") & ", " & drs("TOPIC_ID") & ", " & 0
		           	strsqlvalues = strsqlvalues & ", '" & chkstring(drs("T_SUBJECT"),"archive") & "', '" & chkstring(drs("T_MESSAGE"),"archive")
		           	strsqlvalues = strsqlvalues & "', " & drs("T_AUTHOR") & ", " & drs("T_REPLIES") & ", " & intT_UREPLIES & ", " & drs("T_VIEW_COUNT")
	        	   	strsqlvalues = strsqlvalues & ", '" & drs("T_LAST_POST") & "', '" & drs("T_DATE") & "', " & drs("T_LAST_POSTER")
	           		strsqlvalues = strsqlvalues & ", '" & drs("T_IP") & "', " & drs("T_LAST_POST_AUTHOR") & ", " & intT_LAST_POST_REPLY_ID & ", '" & drs("T_LAST_EDIT")
				strsqlvalues = strsqlvalues & "', " & intT_LAST_EDITBY & ", " & drs("T_STICKY") & ", " & drs("T_SIG") & " "

		       		strsql = "INSERT INTO " & strArchiveTablePrefix & "TOPICS (CAT_ID, FORUM_ID, TOPIC_ID, T_STATUS, T_SUBJECT, T_MESSAGE, T_AUTHOR, T_REPLIES, T_UREPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID, T_LAST_EDIT, T_LAST_EDITBY, T_STICKY, T_SIG)"
				strsql = strsql & " VALUES (" & strsqlvalues & ")"
				'Response.Write strSql
				'Response.End
				my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
				msg = "                      <center><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>All topics older than " & strToDate(fdateolderthan) & " were archived</font></center><br />" & vbNewLine
			else
		       		strsql = "UPDATE " & strArchiveTablePrefix & "TOPICS SET " &_
					"T_STATUS = " & 0 &_
					", T_SUBJECT = '" & chkstring(drs("T_SUBJECT"),"archive") & "'" &_
					", T_MESSAGE = '" & chkstring(drs("T_MESSAGE"),"archive") & "'" &_
					", T_REPLIES = T_REPLIES + " & drs("T_REPLIES") &_
					", T_UREPLIES = T_UREPLIES + " & intT_UREPLIEScnt &_
					", T_VIEW_COUNT = T_VIEW_COUNT + " & drs("T_VIEW_COUNT") &_
					", T_LAST_POST = '" & drs("T_LAST_POST") & "'" &_ 
					", T_LAST_POST_AUTHOR = " & drs("T_LAST_POST_AUTHOR") &_
					", T_LAST_POST_REPLY_ID = " & intT_LAST_POST_REPLY_ID & _
					", T_LAST_EDIT = '" & drs("T_LAST_EDIT") & "'" & _
					", T_LAST_EDITBY = " & intT_LAST_EDITBY & _
					", T_STICKY = " & drs("T_STICKY") & _
					", T_SIG = " & drs("T_SIG") & _
					" WHERE TOPIC_ID = " & drs("TOPIC_ID")
 	            		response.write("                      <font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>." & vbNewLine)
				my_conn.execute(strsql),,adCmdText + adExecuteNoRecords

				msg = "                      <br /><center>Topic exists, Stats Updated......</center></font>" & vbNewLine
			end if

		        Response.Write msg
			
			'#### Delete originals
			if i > 100 then
				i = 0
				response.write("                      <br />" & vbNewLine)
			end if
			i = i + 1

			'## Forum_SQL - Delete any subscriptions to this topic
			strSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS "
			strSql = strSql & " WHERE TOPIC_ID = " & drs("TOPIC_ID")
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
           drs.movenext
	Loop
	drs.close
	strSql = "DELETE FROM " & strTablePrefix & "TOPICS WHERE T_LAST_POST < '" & fdateolderthan & "' " & fIDSQL
	strSqL = strSqL & " AND T_ARCHIVE_FLAG <> 0 "
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
    End if
    Call subdoupdates()
End Sub

Sub subdeletestuff(fstrid)
	Dim fIDSQL
	rqID = request("id")
	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then 
			fIDSQL = " WHERE FORUM_ID=" & rqID
		else
			fIDSQL = ""
		end if
		err.clear
	else
		fIDSQL = " WHERE FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0
	
	strsql = "DELETE FROM " & strTablePrefix & "TOPICS " & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	strsql = "DELETE FROM " & strTablePrefix & "REPLY " & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	
	'#### Update FORUM last delete posts date
	strsql = "UPDATE " & strTablePrefix & "FORUM SET F_L_DELETE= '" & DateToStr(now()) & "'"
	strsql = strsql & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	
	Call subdoupdates()
End Sub

Sub subdoupdates()
	'#### create FORUM_ID clause
	rqID = request("id")
    	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then 
			fIDSQL = " AND " & strTablePrefix & "FORUM.FORUM_ID=" & rqID
			fIDSQL2 = " WHERE " & strTablePrefix & "TOPICS.FORUM_ID=" & rqID
		else
			fIDSQL = ""
			fIDSQL2 = ""
		end if
		err.clear
	else
		fIDSQL = " AND " & strTablePrefix & "FORUM.FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		fIDSQL2 = " WHERE " & strTablePrefix & "TOPICS.FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0

	response.write	"<table align=""center"" border=""0"">" & vbNewLine & _
			"<tr>" & vbNewLine & _
			"<td align=""center"" colspan=""2""><p><b><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>Updating Counts</font></b><br /></td>" & vbNewLine & _
			"</tr>" & vbNewLine & _
			"<tr>" & vbNewLine & _
			"<td align=""right"" valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>Topics:</font></td>" & vbNewLine & _
			"<td valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>"

	set rs = Server.CreateObject("ADODB.Recordset")
	set rs1 = Server.CreateObject("ADODB.Recordset")

	'## Forum_SQL - Get contents of the Forum table related to counting
	strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 " & fIDSQL

	rs.Open strSql, my_Conn
	if not(rs.EOF or rs.BOF) then
		rs.MoveFirst
		i = 0 

		do until rs.EOF
			i = i + 1
			'## Forum_SQL - count total number of topics in each forum in Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")

			set rs1 = my_Conn.Execute( strSql)
			if rs1.EOF or rs1.BOF then
				intF_TOPICS = 0
			else
				intF_TOPICS = rs1("cnt")
			end if
			rs1.Close

			'## Forum_SQL - count total number of archived topics in each forum in A_Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")

			set rs1 = my_Conn.Execute( strSql)
			if rs1.EOF or rs1.BOF then
				intF_A_TOPICS = 0
			else
				intF_A_TOPICS = rs1("cnt")
			end if
			rs1.Close

			strSql = "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET F_TOPICS = " & intF_TOPICS
			strSql = strSql & " , F_A_TOPICS = " & intF_A_TOPICS
			strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			rs.MoveNext
			Response.Write "."
			if i = 80 then 
				Response.Write "<br />"
				i = 0
			end if
		loop
	end if
	rs.Close

	Response.Write	"</font></td>" & vbNewLine & _
			"</tr>" & vbNewLine & _
			"<tr>" & vbNewLine & _
			"<td align=""right"" valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>Topic Replies:</font></td>" & vbNewLine & _
			"<td valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>"

	'## Forum_SQL
	strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS" & fIDSQL2

	rs.Open strSql, my_Conn
	i = 0 

	do until rs.EOF
		i = i + 1

		'## Forum_SQL - count total number of replies in Topics table
		strSql = "SELECT count(REPLY_ID) AS cnt "
		strSql = strSql & " FROM " & strTablePrefix & "REPLY "
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")

		rs1.Open strSql, my_Conn
		if rs1.EOF or rs1.BOF or (rs1("cnt") = 0) then
			intT_REPLIES = 0
		else
			intT_REPLIES = rs1("cnt")
		end if
	
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")

		my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

		rs1.Close
		rs.MoveNext
		Response.Write "."
		if i = 80 then 
			Response.Write "<br />"
			i = 0
		end if
	loop
	rs.Close

	Response.Write 	"</font></td>" & vbNewline & _
			"</tr>" & vbNewline & _
			"<tr>" & vbNewline & _
			"<td align=""right"" valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>Forum Replies:</font></td>" & vbNewline & _
			"<td valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>"

	'## Forum_SQL - Get values from Forum table needed to count replies
	strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	rs.Open strSql, my_Conn, adOpenDynamic, adLockOptimistic, adCmdText

	do until rs.EOF
		'## Forum_SQL - Count total number of Replies
		strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
		strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & rs("FORUM_ID")

		rs1.Open strSql, my_Conn

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

		rs1.Close

		'## Forum_SQL - Count total number of Archived Replies
		strSql = "SELECT Sum(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS SumOfT_A_REPLIES, Count(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS cnt "
		strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
		strSql = strSql & " WHERE " & strTablePrefix & "A_TOPICS.FORUM_ID = " & rs("FORUM_ID")
	
		rs1.Open strSql, my_Conn

		if rs1.EOF or rs1.BOF then
			intF_A_COUNT = 0
			intF_A_TOPICS = 0
		else
			if IsNull(rs1("SumOfT_A_REPLIES")) then
				intF_A_COUNT = rs1("cnt")
			else
				intF_A_COUNT = CLng(rs1("cnt")) + CLng(rs1("SumOfT_A_REPLIES"))
			end if
			intF_A_TOPICS = rs1("cnt") 
		end if
		if IsNull(intF_A_COUNT) then intF_A_COUNT = 0 
		if IsNull(intF_A_TOPICS) then intF_A_TOPICS = 0 

		rs1.Close

		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET F_COUNT = " & intF_COUNT
		strSql = strSql & ",  F_TOPICS = " & intF_TOPICS
		strSql = strSql & ",  F_A_COUNT = " & intF_A_COUNT
		strSql = strSql & ",  F_A_TOPICS = " & intF_A_TOPICS
		strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
	
		my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

		rs.MoveNext
		Response.Write "."
		if i = 80 then 
			Response.Write "<br />" & vbNewline
			i = 0
		end if	
	loop
	rs.Close

	Response.Write	"</font></td>" & vbNewline & _
			"</tr>" & vbNewline & _
			"<tr>" & vbNewline & _
			"<td align=""right"" valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>Totals:</font></td>" & vbNewline & _
			"<td valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>"
	'## Forum_SQL - Total of Topics
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
	strSql = strSql & " AS SumOfF_TOPICS "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	rs.Open strSql, my_Conn
	
	if IsNull(RS("SumOfF_TOPICS")) then
		Response.Write "Total Topics: 0<br />" & vbNewline
		strSumOfF_TOPICS = 0
	else
		Response.Write "Total Topics: " & RS("SumOfF_TOPICS") & "<br />" & vbNewline
		strSumOfF_TOPICS = rs("SumOfF_TOPICS")
	end if
	rs.Close

	'## Forum_SQL - Total of Archived Topics
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_A_TOPICS) "
	strSql = strSql & " AS SumOfF_A_TOPICS "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	rs.Open strSql, my_Conn

	if IsNull(RS("SumOfF_A_TOPICS")) then
		Response.Write "Total Archived Topics: 0<br />" & vbNewline
		strSumOfF_A_TOPICS = 0
	else
		Response.Write "Total Archived Topics: " & RS("SumOfF_A_TOPICS") & "<br />" & vbNewline
		strSumOfF_A_TOPICS = rs("SumOfF_A_TOPICS")
	end if
	rs.Close
	
	'## Forum_SQL - Total all the replies for each topic
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
	strSql = strSql & " AS SumOfF_COUNT "
	strSql = strSql & ", Sum(" & strTablePrefix & "FORUM.F_A_COUNT) "
	strSql = strSql & " AS SumOfF_A_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = my_Conn.Execute (strSql)

	if rs("SumOfF_COUNT") <> "" then
		Response.Write "Total Posts: " & RS("SumOfF_COUNT") & "<br />" & vbNewline
		strSumOfF_COUNT = rs("SumOfF_COUNT")
	else
		Response.Write "Total Posts: 0<br />" & vbNewline
		strSumOfF_COUNT = "0"
	end if

	if rs("SumOfF_A_COUNT") <> "" then
		Response.Write "Total Archived Posts: " & RS("SumOfF_A_COUNT") & "<br />" & vbNewline
		strSumOfF_A_COUNT = rs("SumOfF_A_COUNT")
	else
		Response.Write "Total Archived Posts: 0<br />" & vbNewline
		strSumOfF_A_COUNT = "0"
	end if

	set rs = nothing
	'## Forum_SQL - Write totals to the Totals table
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET T_COUNT = " & strSumOfF_TOPICS
	strSql = strSql & ", P_COUNT = " & strSumOfF_COUNT
	strSql = strSql & ", T_A_COUNT = " & strSumOfF_A_TOPICS
	strSql = strSql & ", P_A_COUNT = " & strSumOfF_A_COUNT

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	Response.Write	"</font></td>" & vbNewline & _
			"</tr>" & vbNewLine & _
			"<tr>" & vbNewLine & _
			"<td align=""center"" colspan=""2"">&nbsp;<br /><b><font face=""" & strDefaultFontFace & """ size=""" & strfooterFontSize & """ color=""" & strForumFontColor & """>Count Update Complete</font></b></font></td>" & vbNewline & _
			"</tr>" & vbNewLine & _
			"</table>"
	set rs = nothing
	set rs1 = nothing
End Sub

WriteFooter
Response.End
%>
