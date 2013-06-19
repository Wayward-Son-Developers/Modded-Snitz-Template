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

' get the data
'strSql = "SELECT CAT_ID COUNT(*) AS PostCount"
'strSql = strSql & " FROM " & strTablePrefix & "TOPICS"
'strSql = strSql & " WHERE T_STATUS > 1"
'strSql = strSql & " GROUP BY CAT_ID"

' CheckforUnmoderatedPosts - This function will check for unmoderated posts by
'							 Board, Category or Forum
function CheckForUnmoderatedPosts(CType, CatID, ForumID, TopicID)
	Dim PostCount
	PostCount = 0
	if strModeration > 0 then
		' Check the Topics Table first
		strSql = "Select Count(*) as PostCount"
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS T"
		if CType = "CAT" then
			strSql = strSql & " WHERE T.CAT_ID = " & CatID & " AND T.T_STATUS > 1 "
		elseif CType = "FORUM" then
			strSql = strSql & " WHERE T.FORUM_ID = " & ForumID & " AND T.T_STATUS > 1 "
		elseif CType = "TOPIC" then
			strSql = strSql & " WHERE T.TOPIC_ID = " & TopicID & " AND T.T_STATUS > 1 "
		elseif CType = "POSTAUTHOR" then
			strSql = strSql & " WHERE T.T_AUTHOR = " & MemberID & " AND T.T_STATUS > 1 AND T.TOPIC_ID = " & TopicID
		end if
		if CType = "BOARD" then
			strSql = strSql & ", " & strTablePrefix & "CATEGORY C"
			strSql = strSql & ", " & strtablePrefix & "FORUM F"
			' This line makes sure that moderation is still set in the Category
			strSql = strSql & " WHERE T.CAT_ID = C.CAT_ID AND C.CAT_MODERATION > 0"
			' This line makes sure that moderation is still set to all posts or topic in the Forum
			strSql = strSql & " AND T.FORUM_ID = F.FORUM_ID AND F.F_MODERATION in (1,2)" & " AND T.T_STATUS > 1 "
		end if
		set rsCheck = my_Conn.Execute(strSql)
		if not rsCheck.EOF then
			PostCount = rsCheck("PostCount")
		else
			PostCount = 0
		end if
		if PostCount = 0 then
			' If no unmoderated posts are found on the topic table, check the replies.....
			strSql = "Select Count(*) as PostCount"
			strSql = strSql & " FROM " & strTablePrefix & "REPLY R"
			if CType = "CAT" then
				strSql = strSql & " WHERE R.CAT_ID = " & CatID & " AND R.R_STATUS > 1 "
			elseif CType = "FORUM" then
				strSql = strSql & " WHERE R.FORUM_ID = " & ForumID & " AND R.R_STATUS > 1 "
			elseif CType = "TOPIC" then
				strSql = strSql & " WHERE R.TOPIC_ID = " & TopicID & " AND R.R_STATUS > 1 "
			elseif cType = "POSTAUTHOR" then
				strSql = strSql & " WHERE R.R_AUTHOR = " & MemberID & " AND R.R_STATUS > 1 AND R.TOPIC_ID = " & TopicID
			end if
			if CType = "BOARD" then
				strSql = strSql & ", " & strTablePrefix & "CATEGORY C"
				strSql = strSql & ", " & strtablePrefix & "FORUM F"
				' This line makes sure that moderation is still set in the Category
				strSql = strSql & " WHERE R.CAT_ID = C.CAT_ID AND C.CAT_MODERATION > 0"
				' This line makes sure that moderation is still set to all posts or reply in the Forum
				strSql = strSql & " AND R.FORUM_ID = F.FORUM_ID AND F.F_MODERATION in (1,3)" & " AND R.R_STATUS > 1 "
			end if
			rsCheck.close
			set rsCheck = my_Conn.Execute(strSql)
			if not rsCheck.EOF then
				PostCount = rsCheck("PostCount")
			else
				PostCount = 0
			end if
		end if
		rsCheck.close
		set rsCheck = nothing
	end if
	CheckForUnModeratedPosts = PostCount
end function
%>
