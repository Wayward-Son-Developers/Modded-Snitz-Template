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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<!--#INCLUDE FILE="inc_func_count.asp" -->
<%
Server.ScriptTimeout = 90
' -- Declare the variables and initialize them with the values from either the querystring (1st time
' -- into the form) or the form (all other times through the form)
' -- Mode - 1 = Approve, 2 = Hold, 3 = Reject
Dim Mode, ModLevel, CatID, ForumID, TopicID, ReplyID, Password, Result, Comments

CatID    = clng("0" & Request("CAT_ID"))
ForumID  = clng("0" & Request("FORUM_ID"))
TopicID  = clng("0" & Request("TOPIC_ID"))
if Request("REPLY_ID") = "X" then
	ReplyID = "X"
else
	ReplyID  = clng("0" & Request("REPLY_ID"))
end if
Comments = trim(Request.Form("COMMENTS"))

' Mode: 1 = Approve, 2 = Hold, 3 = Reject
Mode = Request("MODE")
if Mode = "" then
	Mode = 0
end if

' Set the ModLevel for the operation
if Mode > 0 then
	if CatID = "0" or CatID = "" then
		ModLevel = "BOARD"
	elseif ForumID = "0"  or ForumID = "" then
		ModLevel = "CAT"
	elseif TopicID = "0"  or TopicID = "" then
		ModLevel = "FORUM"
	elseif ReplyId = "0"  or ReplyID = "" then
		ModLevel = "TOPIC"
	elseif ReplyId = "X" then
		ModLevel = "ALLPOSTS"
	else
		ModLevel = "REPLY"
	end if
end if

if mlev = 0 then
	Call FailMessage("<li>You must be logged in to Moderate posts.</li>",False)
elseif Mode = "" or Mode = 0 then
	ModeForm
else
	if ModLevel = "BOARD" or ModLevel = "CAT" then
		if mlev < 4 then
			Response.Write	"<p class=""warning"">Only Admins May "
			if Mode = 1 then
				Result = "Approve "
			elseif Mode = 2 then
				Result = "Hold "
			else
				Result = "Reject "
			end if
			if ModLevel = "BOARD" then
				Result = Result & "all Topics and Replies for the Forum. "
			else
				Result = Result & "the Topics and Replies for this Category. "
			end if
			Response.Write Result & "</p>" & vbNewline
			LoginForm
		elseif Mode = 1 or Mode = 2 then
			Approve_Hold
		else
			Delete
		end if
	else
		' -- Not an admin or moderator.  Can't do...
		if mlev < 4 and chkforumModerator(ForumID, strDBNTUserName) <> "1" then
			Response.Write	"<p class=""warning"">Only Admins and Moderators May "
			if Mode = 1 then
				Result = "Approve "
			elseif Mode = 2 then
				Result = "Hold "
			else
				Result = "Reject "
			end if
			if ModLevel = "FORUM" then
				Result = Result & "all Topics and Replies for the Forum. "
			elseif ModLevel = "TOPIC" then
				Result = Result & "this Topic. "
			elseif ModLevel = "ALLPOSTS" then
				Result = Result & "all Posts for this Topic. "
			else
				Result = Result &  "this Reply. "
			end if
			Response.Write	Result & "</p>" & vbNewline
			LoginForm
		elseif Mode = 1 or Mode = 2 then
			' -- Do the approval/Hold
			Approve_Hold
		else
			Delete
		end if
	end if
end if

WriteFooterShort
Response.End

sub Approve_Hold
	' Loop through the topic table to determine which records need to be updated.
	if ModLevel <> "Reply" then
		strSql = "SELECT T.CAT_ID, "
		strSql = strSql & "T.FORUM_ID, "
		strSql = strSql & "T.TOPIC_ID, "
		strSql = strSql & "T.T_LAST_POST as Post_Date, "
		strSql = strSql & "M.M_NAME, "
		strSql = strSql & "M.MEMBER_ID "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS T, "
		strSql = strSql & strMemberTablePrefix & "MEMBERS M"
		strSql = strSql & " WHERE (T.T_STATUS = 2 OR T.T_STATUS = 3) "
		strSql = strSql & "   AND T.T_AUTHOR = M.MEMBER_ID"
		' Set the appropriate level of moderation based on the passed mode.
		if ModLevel <> "BOARD" then
			if Modlevel = "CAT" then
				strSql = strSql & " AND T.CAT_ID = " & CatID
			elseif Modlevel = "FORUM" then
				strSql = strSql & " AND T.FORUM_ID = " & ForumID
			else
				strSql = strSql & " AND T.TOPIC_ID = " & TopicID
			end if
		end if
		set rsLoop = my_Conn.Execute (strSql)
		if rsLoop.EOF or rsLoop.BOF then
			' Do nothing - No records meet this criteria
		else
			do until rsLoop.EOF
				LoopCatID      = rsLoop("CAT_ID")
				LoopForumID    = rsLoop("FORUM_ID")
				LoopTopicID    = rsLoop("TOPIC_ID")
				LoopMemberID   = rsLoop("MEMBER_ID")
				LoopMemberName = rsLoop("M_NAME")
				LoopPostDate   = rsLoop("POST_DATE")

				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " set T_STATUS = "

				if Mode = 1 then
					StrSql = StrSql & " 1"
					strSql = strSql & " , T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
					strSql = strSql & " , T_LAST_POST_REPLY_ID = " & 0
					LoopPostDate = DateToStr(strForumTimeAdjust)
				else
					StrSql = StrSql & " 3"
				end if
				strSql = strSql & " WHERE CAT_ID = " & LoopCatID
				strSql = strSql & " AND FORUM_ID = " & LoopForumID
				strSql = strSql & " AND TOPIC_ID = " & LoopTopicID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				' -- If approving, make sure to update the appropriate counts..
				if Comments <> "" then
					Send_Comment_Email LoopMemberName, LoopMemberID, LoopCatID, LoopForumID, LoopTopicID, 0
				end if
				if Mode = 1 then
					doPCount
					doTCount
					UpdateForum "Topic", LoopForumID, LoopMemberID, LoopPostDate, LoopTopicID, 0
					UpdateUser LoopMemberID, LoopForumID, LoopPostDate
					ProcessSubscriptions LoopMemberID, LoopCatID, LoopForumID, LoopTopicID, "No"
				end if
				rsLoop.MoveNext
			loop
		end if
		rsLoop.Close
		set rsLoop = nothing
	end if

	' Update the replies if appropriate
	strSql = "SELECT R.CAT_ID, " & _
			 "R.FORUM_ID, " & _
			 "R.TOPIC_ID, " & _
			 "R.REPLY_ID, " & _
			 "R.R_DATE as Post_Date, " & _
			 "M.M_NAME, " & _
			 "M.MEMBER_ID " & _
			 " FROM " & strTablePrefix & "REPLY R, " & _
			 strMemberTablePrefix & "MEMBERS M" & _
			 " WHERE (R.R_STATUS = 2 OR R.R_STATUS = 3) " & _
			 " AND R.R_AUTHOR = M.MEMBER_ID "
	if ModLevel <> "BOARD" then
		if ModLevel = "CAT" then
			strSql = strSql & " AND R.CAT_ID = " & CatID
		elseif ModLevel = "FORUM" then
			strSql = strSql & " AND R.FORUM_ID = " & ForumID
		elseif ModLevel = "TOPIC" or ModLevel = "ALLPOSTS" then
			strSql = strSql & " AND R.TOPIC_ID = " & TopicID
		else
			strSql = strSql & "AND R.REPLY_ID = " & ReplyID
		end if
	end if
	set rsLoop = my_Conn.Execute (strSql)
	if rsLoop.EOF or rsLoop.BOF then
		' Do nothing - No records matching the criteria were found
	else
		do until rsLoop.EOF
			LoopMemberName = rsLoop("M_NAME")
			LoopMemberID   = rsLoop("MEMBER_ID")
			LoopCatID      = rsLoop("CAT_ID")
			LoopForumID    = rsLoop("FORUM_ID")
			LoopTopicID    = rsLoop("TOPIC_ID")
			LoopReplyID    = rsLoop("REPLY_ID")
			LoopPostDate   = rsLoop("POST_DATE")
			StrSql = "UPDATE " & strTablePrefix & "REPLY "
			StrSql = StrSql & " set R_STATUS = "
			if Mode = 1 then
				StrSql = StrSql & " 1"
				strSql = strSql & " , R_LAST_EDIT = '" & DateToStr(strForumTimeAdjust) & "'"
				LoopPostDate = DateToStr(strForumTimeAdjust)
			else
				StrSql = StrSql & " 3"
			end if
			StrSql = StrSql & " WHERE REPLY_ID = " & LoopReplyID
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			if Comments <> "" then
				Send_Comment_Email LoopMemberName, LoopMemberID, LoopCatID, LoopForumID, LoopTopicID, LoopReplyID
			end if
			if Mode = 1 then 
				doPCount
		                UpdateTopic LoopTopicID, LoopMemberID, LoopPostDate, LoopReplyID
		                UpdateForum "Post", LoopForumID, LoopMemberID, LoopPostDate, LoopTopicID, LoopReplyID
		                UpdateUser LoopMemberID, LoopForumID, LoopPostDate
		                ProcessSubscriptions LoopMemberID, LoopCatID, LoopForumID, LoopTopicID, "No"
			end if
			rsLoop.MoveNext
		loop
	end if
	rsLoop.Close
	set rsLoop = nothing

	' ## Build final result message
	if ModLevel = "BOARD" then
		Result = "All Topics and Replies have "
	elseif ModLevel = "CAT" then
		Result = "All Topics and Replies in this Category have "
	elseif ModLevel = "FORUM" then
		Result = "All Topics and Replies in this Forum have "
	elseif ModLevel = "TOPIC"  then
		Result = "This Topic has "
	elseif ModLevel = "ALLPOSTS" then
		Result = "All posts for this topic have "
	else
		Result = "This Reply has "
	end if
	if Mode = 2 then
		Result = Result & " Been Placed on Hold."
	elseif Mode = 3 then 
		Result = Result & " Been Deleted."
	else
		Result = Result & " Been Approved."
	end if
	
	Call PopOkMessage(Result,True)
end sub

sub Delete
	' Loop through the topic table to determine which records need to be updated.
	if ModLevel <> "Reply" then
		strSql = "SELECT T.CAT_ID, "
		strSql = strSql & "T.FORUM_ID, "
		strSql = strSql & "T.TOPIC_ID, "
		strSql = strSql & "T.T_LAST_POST as Post_Date, "
		strSql = strSql & "M.M_NAME, "
		strSql = strSql & "M.MEMBER_ID "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS T, "
		strSql = strSql & strMemberTablePrefix & "MEMBERS M"
		strSql = strSql & " WHERE (T.T_STATUS = 2 OR T.T_STATUS = 3) "
		strSql = strSql & "   AND T.T_AUTHOR = M.MEMBER_ID"
		' Set the appropriate level of moderation based on the passed mode.
		if ModLevel <> "BOARD" then
			if Modlevel = "CAT" then
				strSql = strSql & " AND T.CAT_ID = " & CatID
			elseif Modlevel = "FORUM" then
				strSql = strSql & " AND T.FORUM_ID = " & ForumID
			else
				strSql = strSql & " AND T.TOPIC_ID = " & TopicID
			end if
		end if
		set rsLoop = my_Conn.Execute (strSql)
		if rsLoop.EOF or rsLoop.BOF then
			' Do nothing - No records meet this criteria
		else
			do until rsLoop.EOF
				LoopCatId      = rsLoop("CAT_ID")
				LoopForumID    = rsLoop("FORUM_ID")
				LoopTopicID    = rsLoop("TOPIC_ID")
				LoopMemberName = rsLoop("M_NAME")
				LoopMemberID   = rsLoop("MEMBER_ID")
		            	if Comments <> "" then
					Send_Comment_Email LoopMemberName, LoopMemberID, LoopCatID, LoopForumID, LoopTopicID, 0
				end if

				strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE CAT_ID = " & LoopCatID
				strSql = strSql & " AND FORUM_ID = " & LoopForumID
				strSql = strSql & " AND TOPIC_ID = " & LoopTopicID
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				' -- If approving, make sure to update the appropriate counts..
				rsLoop.MoveNext
			loop
		end if
		rsLoop.Close
		set rsLoop = nothing
	end if

	' Update the replies if appropriate
	strSql = "SELECT R.CAT_ID, " & _
			 "R.FORUM_ID, " & _
			 "R.TOPIC_ID, " & _
			 "R.REPLY_ID, " & _
			 "R.R_STATUS, " & _
			 "R.R_DATE as Post_Date, " & _
			 "M.M_NAME, " & _
			 "M.MEMBER_ID " & _
			 " FROM " & strTablePrefix & "REPLY R, " & strMemberTablePrefix & "MEMBERS M" & _
			 " WHERE (R.R_Status = 2 OR R.R_Status = 3) " & _
			 " AND R.R_AUTHOR = M.MEMBER_ID "
	if ModLevel <> "BOARD" then
		if ModLevel = "CAT" then
			strSql = strSql & " AND R.CAT_ID = " & CatID
		elseif ModLevel = "FORUM" then
			strSql = strSql & " AND R.FORUM_ID = " & ForumID
		elseif ModLevel = "TOPIC" or ModLevel = "ALLPOSTS" then
			strSql = strSql & " AND R.TOPIC_ID = " & TopicID
		else
			strSql = strSql & " AND R.REPLY_ID = " & ReplyID
		end if
	end if
	set rsLoop = my_Conn.Execute (strSql)
	if rsLoop.EOF or rsLoop.BOF then
		' Do nothing - No records matching the criteria were found
	else
		do until rsLoop.EOF
			if Comments <> "" then
				Send_Comment_Email rsLoop("M_NAME"), rsLoop("MEMBER_ID"), rsLoop("CAT_ID"), rsLoop("FORUM_ID"), rsLoop("TOPIC_ID"), rsLoop("REPLY_ID")
			end if
			if rsLoop("R_STATUS") = 2 then
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_UREPLIES = T_UREPLIES - 1 "
				strSql = strSql & " WHERE TOPIC_ID = " & rsLoop("TOPIC_ID")
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			end if
			StrSql = "DELETE FROM " & strTablePrefix & "REPLY "
			StrSql = StrSql & " WHERE REPLY_ID = " & rsLoop("REPLY_ID")
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			rsLoop.MoveNext
		loop
	end if
	rsLoop.Close
	set rsLoop = nothing

	' ## Build final result message
	if ModLevel = "BOARD" then
		Result = "All Topics and Replies have "
	elseif ModLevel = "CAT" then
		Result = "All Topics and Replies in this Category have "
	elseif ModLevel = "FORUM" then
		Result = "All Topics and Replies in this Forum have "
	elseif ModLevel = "TOPIC"  then
		Result = "This Topic has "
	elseif ModLevel = "ALLPOSTS" then
		Result = "All posts for this topic have "
	else
		Result = "This Reply has "
	end if
	if Mode = 2 then
		Result = Result & " Been Placed on Hold."
	elseIf Mode = 3 then
		Result = Result & " Been Deleted."
	else
		Result = Result & " Been Approved."
	end if
	
	Call PopOkMessage(Result,True)
end sub

' ## ModeForm - This is the form which is used to determine exactly what the admin/moderator wants
' ## to do with the posts he is working on.
sub ModeForm
	Response.Write	"<form action=""pop_moderate.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewline & _
					"<input type=""hidden"" name=""REPLY_ID"" value=""" & ReplyID & """>" & vbNewline & _
					"<input type=""hidden"" name=""TOPIC_ID"" value=""" & TopicID & """>" & vbNewline & _
					"<input type=""hidden"" name=""FORUM_ID"" value=""" & ForumID & """>" & vbNewline & _
					"<input type=""hidden"" name=""CAT_ID""   value=""" & CatID & """>" & vbNewline & _
					"<table class=""admin"" width=""100%"">" & vbNewline & _
					"<tr>" & vbNewline & _
					"<td class=""formvalue"">" & vbNewline & _
					"<select name=""Mode"">" & vbNewline & _
					"<option value=""1"" SELECTED>Approve</option>" & vbNewline & _
					"<option value=""2"">Hold</option>" & vbNewline & _
					"<option value=""3"">Delete</option>" & vbNewline & _
					"</select>" & vbNewline 
	If ModLevel = "TOPIC" or ModLevel = "REPLY" then
		Response.Write " this post" & vbNewline
	Else
		Response.Write " these posts" & vbNewline
	End if
	Response.Write	"</td>" & vbNewline & _
					"</tr>" & vbNewline
	if strEmail = 1 then
		Response.Write	"<tr>" & vbNewline & _
						"<td class=""formvalue"">Comments:" & vbNewline & _ 
						"<textarea name=""Comments"" cols=""45"" rows=""6"" wrap=""VIRTUAL""></textarea><br />" & vbNewline & _
						"<span face=""smtext"">The comments you type herewill be mailed to the author of the topic(s)</span></td>" & vbNewline & _
						"</tr>" & vbNewline
	end if
	Response.Write	"<tr>" & vbNewline & _
					"<td class=""options""><button type=""Submit"" id=""Submit1"" name=""Submit1"">Send</button></td>" & vbNewline & _
					"</tr>" & vbNewline & _
					"</table>"  & vbNewline & _
					"</form>" & vbNewline
end Sub

' ## UpdateForum - This will update the forum table by adding to the total
' ##               posts (and total topics if appropriate),
' ##               and will also update the last forum post date and poster if
' ##               appropriate.
sub UpdateForum(UpdateType, ForumID, MemberID, PostDate, TopicID, ReplyID)
	dim UpdateLastPost
	' -- Check the last date/time to see if they need updating.
	strSql = " SELECT F_LAST_POST "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE FORUM_ID = " & ForumID
	set RsCheck = my_Conn.Execute (strSql)
	if rsCheck("F_LAST_POST") < PostDate then
		UpdateLastPost = "Y"
	end if
	rsCheck.Close
	set rsCheck = nothing

	strSql = "UPDATE " & strTablePrefix & "FORUM "
	strSql = strSql & " SET F_COUNT = F_COUNT + 1 "
	if UpdateType = "Topic" then strSql = strSql & ", F_TOPICS = F_TOPICS + 1 "
	if UpdateLastPost = "Y" then
		strSql = strSql & ", F_LAST_POST = '" & PostDate & "'"
		strSql = strSql & ", F_LAST_POST_AUTHOR = " & MemberID
		strSql = strSql & ", F_LAST_POST_TOPIC_ID = " & TopicID
		strSql = strSql & ", F_LAST_POST_REPLY_ID = " & ReplyID
	end if
	strSql = strSql & " WHERE FORUM_ID = " & ForumID
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

' ## UpdateTopic - This will update the T_REPLIES field (and T_LAST_POST , T_LAST_POSTER & T_UREPLIES if applicable)
' ##               for the appropriate topic
sub UpdateTopic(TopicID, MemberID, PostDate, ReplyID)
	dim UpdateLastPost
	' -- Check the last date/time to see if they need updating.
	strSql = " SELECT T_LAST_POST, T_UREPLIES "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE TOPIC_ID = " & TopicID
	set RsCheck = my_Conn.Execute (strSql)
	if rsCheck("T_LAST_POST") < PostDate then
		UpdateLastPost = "Y"
	end if
	if rsCheck("T_UREPLIES") > 0 then
		UpdateUReplies = "Y"
	end if
	rsCheck.Close
	set rsCheck = nothing

	strSql = "UPDATE " & strTablePrefix & "TOPICS "
	strSql = strSql & " SET T_REPLIES = T_REPLIES + 1 "
	if UpdateLastPost = "Y" then
		strSql = strSql & ", T_LAST_POST = '" & PostDate & "'"
		strSql = strSql & ", T_LAST_POST_AUTHOR = " & MemberID
		strSql = strSql & ", T_LAST_POST_REPLY_ID = " & ReplyID
	end if
	if UpdateUReplies = "Y" then
		strSql = strSql & ", T_UREPLIES = T_UREPLIES - 1 "
	end if

	strSQL = strSQL & " WHERE TOPIC_ID = " & TopicID
	'Response.Write "strSql = " & strSql
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

' ## UpdateUser - This will update the members table by adding to the total
' ##              posts (and total topics if appropriate), and will also update
' ##              the last forum post date and poster if appropriate.
sub UpdateUser(MemberID, LForumID, PostDate)
	dim UpdateLastPost
	' -- Check to see if this post is the newest one for the member...
	set rsCheck = my_Conn.Execute("SELECT M_LASTPOSTDATE FROM " & strMemberTablePrefix & "MEMBERS WHERE MEMBER_ID = " & MemberID)
	if rsCheck("M_LASTPOSTDATE") < PostDate then
		UpdateLastPost = "Y"
	end if
	rsCheck.Close
	set rsCheck = nothing

	set rsFCountMP = my_Conn.Execute("SELECT F_COUNT_M_POSTS FROM " & strTablePrefix & "FORUM WHERE FORUM_ID = " & LForumID)
	ForumCountMPosts = rsFCountMP("F_COUNT_M_POSTS")
	rsFCountMP.close
	set rsFCountMP = nothing

	if UpdateLastPost = "Y" then
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_LASTPOSTDATE = '" & PostDate & "'"
		strSql = strSql & " WHERE MEMBER_ID = " & MemberID
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	end if

	if ForumCountMPosts <> 0 then
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_POSTS = (M_POSTS + 1)"
		strSql = strSql & " WHERE MEMBER_ID = " & MemberID
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	end if
end sub

'## Send_Comment_Email - This sub will send and e-mail to the poster and tell them what the moderator
'##                      or Admin did with their posts.
sub Send_Comment_Email (MemberName, pMemberID, CatID, ForumID, TopicID, ReplyID)

	' -- Get the Admin/Moderator MemberID
	AdminModeratorID = MemberID
	' -- Get the Admin/Moderator Name
	AdminModeratorName = strDBNTUserName
	' -- Get the Admin/Moderator Email
	strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS" & _
	         " WHERE MEMBER_ID = " & AdminModeratorID
	set rsSub = my_Conn.Execute (strSql)
	if rsSub.EOF or rsSub.BOF then
		exit sub
	else
		AdminModeratorEmail = rsSub("M_EMAIL")
	end if

	' -- Get the Category Name and Forum Name
	strSql = "SELECT C.CAT_NAME, F.F_SUBJECT " & _
	         " FROM " & strTablePrefix & "CATEGORY C, " & strTablePrefix & "FORUM F" & _
	         " WHERE C.CAT_ID = " & CatID & " AND F.FORUM_ID = " & ForumID
	set rsSub = my_Conn.Execute (strSql)
	if RsSub.Eof or RsSub.BOF then
 		' Do Nothing -- Should never happen
	else
		CatName = rsSub("CAT_NAME")
		ForumName = rsSub("F_SUBJECT")
	end if

	' -- Get the topic title
	strSql = "SELECT T.T_SUBJECT FROM " & strTablePrefix & "TOPICS T" & _
	         " WHERE T.TOPIC_ID = " & TopicId
	set rsSub = my_Conn.Execute (strSql)
	if rsSub.EOF or rsSub.BOF then
		TopicName = ""
	else
		TopicName = rsSub("T_SUBJECT")
	end if
	rsSub.Close
	set rsSub = Nothing

	strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS" & _
	         " WHERE MEMBER_ID = " & pMemberID
	set rsSub = my_Conn.Execute (strSql)
	if rsSub.EOF or rsSub.BOF then
		exit sub
	else
		MemberEmail = rsSub("M_EMAIL")
	end if

	strRecipientsName = MemberName
	strRecipients = MemberEmail
	strSubject = strForumTitle & " - Your post "
	if Mode = 1 then
		strSubject = strSubject & "has been approved "
	elseif Mode = 2 then
		strSubject = strSubject & "has been placed on hold "
	else
		strSubject = strSubject & "has been rejected "
	end if
	strMessage = "Hello " & MemberName & "." & vbNewline & vbNewline & _
	             " You made a " 
	if Reply = 0 then
		strMessage = strMessage & "post "
	else
		strMessage = strMessage & "reply to the post "
	end if
	strMessage = strMessage & "in the " & ForumName & " forum entitled " & _
	             TopicName & ".  " & AdminModeratorName & " has decided to " 
	if Mode = 1 then
		strMessage = strMessage & "approve your post "
	elseif Mode = 2 then
		strMessage = strMessage & "place your post on hold "
	else
		strMessage = strMessage & "reject your post "
	end if
	strMessage = strMessage & " for the following reason: " & vbNewline & vbNewline & _
	             Comments & vbNewline & vbNewline & _
	             "If you have any questions, please contact " & AdminModeratorName & _
	             " at " & AdminModeratorEmail
	%><!--#INCLUDE FILE="inc_mail.asp" --><%
end sub
%>
