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

sub ProcessSubscriptions (pMemberId, CatID, ForumId, TopicId, Moderation)
	' DEM --> Added line to ignore the moderator/admin since they would be approving the post if
	'         ThisMemberId & MemberID are different....
	ThisMemberID = MemberID
		
	' -- If subscription is not allowed or e-mail is not turned on, exit
	if strSubscription = 0 or strEmail = 0 then
		exit sub
	end if
	StrSql = "SELECT C.CAT_SUBSCRIPTION, C.CAT_NAME, F.F_SUBJECT, F.F_SUBSCRIPTION, " & _
			 "       T.T_SUBJECT, M.M_NAME " & _
             " FROM " & strTablePrefix & "CATEGORY C, " & _
			 "      " & strTablePrefix & "FORUM F, " & _
			 "      " & strTablePrefix & "TOPICS T, " & _
			 "      " & strMemberTablePrefix & "MEMBERS M " & _
             " WHERE C.CAT_ID    = " & CatID   & " AND F.FORUM_ID  = " & ForumID & _
			 "   AND T.TOPIC_ID  = " & TopicID & " AND M.MEMBER_ID = " & pMemberID
	Set rsSub = Server.CreateObject("ADODB.Recordset")
	rsSub.open strSql, my_Conn

	' -- If No record is found, exit sub
	if RsSub.Eof or RsSub.BOF then
		rsSub.close
		set rsSub = nothing
		exit sub
	else
		' Pull the data from the recordset
		allSubsData = rsSub.GetRows(adGetRowsRest)
		SubCnt = UBound(allSubsData,2)
	end if		
	rsSub.close
	set rsSub = nothing
		
	CatSubscription   = allSubsData(0, 0)
	CatName           = allSubsData(1, 0)
	ForumName         = allSubsData(2, 0)
	ForumSubscription = allSubsData(3, 0)
	TopicName         = allSubsData(4, 0)
	MemberName        = allSubsData(5, 0)
	' -- If no subscriptions are allowed for the category or forum, exit sub
	if CatSubscription = 0 or ForumSubscription = 0 then
		exit sub
	end if

	' -- Set highest subscription level to check for...
	' strSubscription   1 = whole board,    2 = by category, 3 = by forum, 4 = by topic
	' CatSubscription   1 = whole category, 2 = by forum,    3 = by topic
	' ForumSubscription 1 = whole forum,    2 = by topic
	If strSubscription = 4 or CatSubscription = 3 or ForumSubscription = 2 then
		SubLevel = "TOPIC"
	Elseif strSubscription = 3 then
		SubLevel = "FORUM"
	ElseIf CatSubscription > 1 then
		SubLevel = "FORUM"
	Elseif StrSubscription > 1 then
		SubLevel = "CATEGORY"
	Else
		SubLevel = "ALL"
	End if

	'## Emails all users who wish to receive a mail if a topic or reply has been made.  This sub will
	'## check for subscriptions based on the topic, forum, category and across the board.  It will 
	'## ignore the posting member.

	if Moderation <> "No" then
		strSql = "SELECT MOD_ID from " & strTablePrefix & "MODERATOR"
		Set modCheck = Server.CreateObject("ADODB.Recordset")
		modCheck.open strSql, my_Conn
		if modCheck.EOF or modCheck.BOF then
			strUniqueModID = "none"
		else
			strUniqueModID = modCheck("Mod_ID")
		end if
		modCheck.Close
		set modCheck = nothing
	else
		strUniqueModID = "none"
	end if

	strSql = "SELECT S.MEMBER_ID, S.CAT_ID, S.FORUM_ID, S.TOPIC_ID, M.M_NAME, M.M_EMAIL " & _
                 " FROM " & strTablePrefix & "SUBSCRIPTIONS S, " & strMemberTablePrefix & "MEMBERS M"
	if Moderation <> "No" and strUniqueModID <> "none" then
		strSql = strSql & ", " & strTablePrefix & "MODERATOR Mo"
	end if
	' -- The author nor the Moderator need to get notification on this topic....
	strSql = strSql & " WHERE S.MEMBER_ID <> " & pMemberID & _
                          " AND S.MEMBER_ID <> " & ThisMemberID & _                 
                          " AND M.MEMBER_ID = S.MEMBER_ID" & _
                          " AND M.M_STATUS <> 0" & _
                          " AND (S.TOPIC_ID = " & TopicId   ' Topic specific subscriptions...

	' -- Check for Subscriptions against the Forum
	if SubLevel <> "TOPIC" then
		StrSql = StrSql & " OR (S.CAT_ID = " & CatID & " AND S.FORUM_ID = " & ForumID & " AND S.TOPIC_ID = 0)"
	end if
	' -- Check for Subscriptions against the Category
	if SubLevel = "CATEGORY" or SubLevel = "ALL" then
		StrSql = StrSql & " OR (S.CAT_ID = " & CatID & " AND S.FORUM_ID = 0 AND S.TOPIC_ID = 0)"
	end if
	' -- Check for Subscriptions against the Board
	if SubLevel = "ALL" then
		StrSql = StrSql & " OR (S.CAT_ID = 0 AND S.FORUM_ID = 0 AND S.TOPIC_ID = 0)"
	end if
	strSql = strSql & ")"

	if Moderation <> "No" then
		StrSql = StrSql & " AND ((M.M_LEVEL = 3"
		if strUniqueModID = "none" then
			StrSql = StrSql & "))"
		else 
			StrSql = StrSql & " AND Mo.MOD_ID = " & strUniqueModID & ") OR (M.M_LEVEL = 2 AND S.MEMBER_ID = Mo.MEMBER_ID AND Mo.FORUM_ID = " & ForumId & "))"
		end if
	end if

	set rsLoop = Server.CreateObject("ADODB.Recordset") : rsLoop.open strSql, my_Conn
	if rsLoop.EOF or rsLoop.BOF then
		rsLoop.close : set rsLoop = nothing : Exit Sub ' No subscriptions, exit....
	else
		' Pull the data from the recordset
		allLoopData = rsLoop.GetRows(adGetRowsRest) : LoopCount = UBound(allLoopData,2)
		rsLoop.close : set rsLoop = nothing

		for iSub = 0 to LoopCount
			LoopMemberID    = allLoopData(0, iSub)
			LoopCatID       = allLoopData(1, iSub)
			LoopForumID     = allLoopData(2, iSub)
			LoopTopicID     = allLoopData(3, iSub)
			LoopMemberName  = allLoopData(4, iSub)
			LoopMemberEmail = allLoopData(5, iSub)
			if chkForumAccess(ForumID, LoopMemberID, false) <> FALSE then
				strRecipientsName = LoopMemberName
				strRecipients = LoopMemberEmail
				strMessage = "Hello " & LoopMemberName & vbNewline & vbNewline
				' ## Send the appropriate message depending on the subscription.
				if LoopCatID > 0 then
					if LoopForumID > 0 then
						if LoopTopicID > 0 then
							strSubject = strForumTitle & " - Reply to a posting"
							strMessage = strMessage & MemberName & " has replied to a topic on " & strForumTitle & " that you requested notification to. "
						else
							strSubject = strForumTitle & " - New posting"
							strMessage = strMessage & MemberName & " has posted to the forum '" & ForumName & "' at " & strForumTitle & " that you requested notification on. "
						end if
					else
						strSubject = strForumTitle & " - New posting"
						strMessage = strMessage & MemberName & " has posted to the category '" & CatName & "' at " & StrForumTitle & " that you requested notification on. "
					end if
				else
					strSubject = strForumTitle & " - New posting"
					strMessage = strMessage & MemberName & " has posted to the " & strForumTitle & " board that you requested notification on. "
				end if
				strMessage = strMessage & "Regarding the subject - " & TopicName & "." & vbNewline & vbNewline
				strMessage = strMessage & "You can view the posting at " & strForumURL & "topic.asp?whichpage=-1&TOPIC_ID=" & TopicId & vbNewline
%>
			<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			end if
		next
	end if
end sub

' PullSubscriptions - will return a list of the subcriptions that exist for a member
Function PullSubscriptions(sCatID, sForumID, sTopicID)
	' -- if subscriptions or e-mail are not turned on, or the person is not logged in, exit...
	If strSubscription = "0" or lcase(strEmail) <> "1" or mlev = 0 then
		PullSubscriptions = ""	:	Exit Function
	End if

	' -- declare the variables used in this function
	Dim BoardSubs, CatSubs, ForumSubs, TopicSubs, rsSub, SubCnt, allSubData, iSub
	Dim SubCatID, SubForumID, SubTopicID
	' -- build the appropriate sql statement...
	subStrSQL = "SELECT CAT_ID, FORUM_ID, TOPIC_ID " & _
			 "  FROM " & strTablePrefix & "SUBSCRIPTIONS" & _
			 " WHERE MEMBER_ID = " & MemberID
	' GetCheck will return the correct SQL statement for the optional parameters....
	subStrSQL = subStrSQL & GetCheck("CAT_ID",   Clng(sCatID))
	subStrSQL = subStrSQL & GetCheck("FORUM_ID", Clng(sForumID))
	subStrSQL = subStrSQL & GetCheck("TOPIC_ID", Clng(sTopicID))
	
	' -- execute the sql statement...
	'Response.Write substrSql
	'Response.End
	Set rsSub = Server.CreateObject("ADODB.Recordset")
	rsSub.open subStrSQL, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rsSub.EOF or rsSub.BOF then
		' If none found, exit
		SubCnt = ""
		PullSubscriptions = ""
	else
		' Pull the data from the recordset
		allSubData = rsSub.GetRows(adGetRowsRest)
		SubCnt = UBound(allSubData,2)
	end if
	rsSub.Close
	set rsSub = Nothing

	if SubCnt = "" then
		' If none found, exit
		PullSubscriptions = ""
	else
		BoardSubs = "N"
		CatSubs = 0
		ForumSubs = 0
		TopicSubs = 0

		for iSub = 0 to SubCnt
			SubCatID   = allSubData(0, iSub)
			SubForumID = allSubData(1, iSub)
			SubTopicID = allSubData(2, iSub)
			If SubCatID = 0 then
				BoardSubs = "Y"
			Elseif SubForumID = 0 then
				If CatSubs > "" then CatSubs = CatSubs & ","
				CatSubs = CatSubs & SubCatID
			Elseif SubTopicID = 0 then
				If ForumSubs > "" then ForumSubs = ForumSubs & ","
				ForumSubs = ForumSubs & SubForumID
			Else
				If TopicSubs > "" then TopicSubs = TopicSubs & ","
				TopicSubs = TopicSubs & SubTopicID
			End If
		next
		PullSubscriptions = BoardSubs & ";" & CatSubs & ";" & ForumSubs & ";" & TopicSubs
	end if
End Function

' GetCheck standardizes the handling of optional parameters in PullSubscriptions
Function GetCheck(ObjectName, ObjectID)
	If ObjectID > 0 then
		GetCheck = " AND " & ObjectName & " = " & ObjectID
	Elseif ObjectID = -99 then
		GetCheck = " AND " & ObjectName & " = 0"
	Else
		GetCheck = ""
	End If
End Function

' Displays the appropriate link, icon and message(if appropriate) for subscriptions...
Function ShowSubLink (SubOption, CatID, ForumID, TopicID, ShowText)
	' -- Declare variables...
	Dim StandardLinkInfo, LinkText, LinkIcon, LinkLevel, LinkParam
	if Instr(Request.ServerVariables("SCRIPT_NAME"),"post.asp") then
		' -- Only show the checkboxes on the post page...
		if SubOption = "S" then
			ShowSubLink = "<input type=""checkbox"" name=""Tnotify"" value=""1"" />&nbsp;Check here to subscribe to this topic."
		else
			ShowSubLink = "<input type=""checkbox"" name=""Tnotify"" value=""0"" />&nbsp;Check here to unsubscribe from this topic."
		end if
	else
		' -- Standard Link
		StandardLinkInfo = "<a href=""Javascript:"
		if SubOption = "U" then
			StandardLinkInfo = StandardLinkInfo & "unsub_confirm"
		else
			StandardLinkInfo = StandardLinkInfo & "openWindow"
		end if
		StandardLinkInfo = StandardLinkInfo & "('pop_subscription.asp?SUBSCRIBE=" & SubOption & "&MEMBER_ID=" & MemberID & "&LEVEL="
		' -- Get appropriate text and icon to display
		LinkParam = ""
		if CatID = 0 then
			LinkLevel = "BOARD"
		else
			LinkParam = "&CAT_ID=" & CatID
			if ForumID = 0 then
				LinkLevel = "CAT"
			else
				LinkParam = LinkParam & "&FORUM_ID=" & ForumID
				if TopicID = 0 then
					LinkLevel = "FORUM"
				else
					LinkLevel = "TOPIC" : LinkParam = LinkParam & "&TOPIC_ID=" & TopicID
				end if
			end if
		end if
		if SubOption = "U" then
			LinkIcon = strIconUnsubscribe
			select case LinkLevel
				case "BOARD" : LinkText = "Unsubscribe from this board"
				case "CAT" : LinkText = "Unsubscribe from this category"
				case "FORUM" : LinkText = "Unsubscribe from this forum"
				case "TOPIC" : LinkText = "Unsubscribe from this topic"
			end select
		else
			LinkIcon = strIconSubscribe
			select case LinkLevel
				case "BOARD" : LinkText = "Subscribe to this board"
				case "CAT" : LinkText = "Subscribe to this category"
				case "FORUM" : LinkText = "Subscribe to this forum"
				case "TOPIC" : LinkText = "Subscribe to this topic"
			end select
		end if
		ShowSubLink = StandardLinkInfo & LinkLevel & LinkParam & "')"">" & getCurrentIcon(LinkIcon, LinkText,"align=""absmiddle""" & dwStatus(LinkText)) & "</a>"
		if ShowText <> "N" then
			ShowSubLink = ShowSubLink & "&nbsp;" & StandardLinkInfo & LinkLevel & LinkParam & "')""" & dwStatus(LinkText) & ">" & LinkText & "</a>" 
		end if
	end if
end function

Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
				"    <!--" & vbNewLine & _
				"    function unsub_confirm(link){" & vbNewLine & _
				"    	var where_to= confirm(""Do you really want to Unsubscribe?"");" & vbNewLine & _
				"       if (where_to== true) {" & vbNewLine & _
				"       	popupWin = window.open(link,'new_page','width=400,height=400')" & vbNewLine & _
				"       }" & vbNewLine & _
				"    }" & vbNewLine & _
				"    //-->" & vbNewLine & _
				"</script>" & vbNewLine
%>
