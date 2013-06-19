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

Function ChkIsNew(dt)
	if Topic_Author = "" then Topic_Author = rs("T_AUTHOR")
	if (Topic_Author = MemberID OR AdminAllowed = 1) AND UnModeratedFPosts > 0 then
		if Topic_Author = MemberID then
			if (CheckForUnModeratedPosts("POSTAUTHOR", Cat_ID, Forum_ID, Topic_ID) > 0) then
				' Do held code
				ChkIsNew = ChkIsNew1()
			else
				' Do normal code
				ChkIsNew = ChkIsNew2(dt)
			end if
		else
			if ((CheckForUnModeratedPosts("TOPIC", Cat_ID, Forum_ID, Topic_ID) > 0) and AdminAllowed = 1) then
				' Do held code
				ChkIsNew = ChkIsNew1()
			else
				' Do normal code
				ChkIsNew = ChkIsNew2(dt)
			end if
		end if
	else
		' Do normal code
		ChkIsNew = ChkIsNew2(dt)
	end if
End function

Function ChkIsNew1()
	if Topic_Status <> 3 then
		UnApprovedFound = "Y"
		ChkIsNew1 = getCurrentIcon(strIconFolderUnmoderated,"Post(s) Awaiting Approval","hspace=""0""")
	else
		HeldFound = "Y"
		ChkIsNew1 = getCurrentIcon(strIconFolderHold,"Post is on Hold","hspace=""0""")
	end if
End Function

Function ChkIsNew2(dt)
	if Topic_Replies = "" then Topic_Replies = rs("T_REPLIES")
	if dt > Session(strCookieURL & "last_here_date") then
		if Topic_Replies >= intHotTopicNum and lcase(strHotTopic) = "1" Then
			ChkIsNew2 = getCurrentIcon(strIconFolderNewHot,"Hot Topic","hspace=""0""")
		else
			ChkIsNew2 = getCurrentIcon(strIconFolderNew,"New Topic","hspace=""0""")
		end if
	elseif Topic_Replies >= intHotTopicNum and lcase(strHotTopic) = "1" Then
		ChkIsNew2 = getCurrentIcon(strIconFolderHot,"Hot Topic","hspace=""0""")
	else
		ChkIsNew2 = getCurrentIcon(strIconFolder,"","hspace=""0""")
	end if
End Function
%>
