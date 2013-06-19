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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_chknew.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<% 
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts

if mLev < 3 then Response.Redirect("default.asp")

Dim ArchiveView
if request("ARCHIVE") = "true" then
	strActivePrefix = strArchiveTablePrefix
	ArchiveView = "true"
	ArchiveLink = "ARCHIVE=true&"
else
	strActivePrefix = strTablePrefix
	ArchiveView = ""
	ArchiveLink = ""
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

'## Forum_SQL - Find all records with the search criteria in them
strSql = "SELECT DISTINCT C.CAT_STATUS, C.CAT_NAME, C.CAT_ORDER"
strSql = strSql & ", F.F_ORDER, F.FORUM_ID, F.F_SUBJECT, F.CAT_ID, F.F_PRIVATEFORUMS"
strSql = strSql & ", F.F_PASSWORD_NEW, F.F_STATUS"
strSql = strSql & ", T.TOPIC_ID, T.T_AUTHOR, T.T_SUBJECT, T.T_STATUS, T.T_LAST_POST"
strSql = strSql & ", T.T_LAST_POST_AUTHOR, T.T_REPLIES, T.T_UREPLIES, T.T_VIEW_COUNT"
strSql = strSql & ", M.MEMBER_ID, M.M_NAME, MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "

strSql2 = " FROM ((((" & strTablePrefix & "FORUM F LEFT JOIN " & strActivePrefix & "TOPICS T"
strSql2 = strSql2 & " ON F.FORUM_ID = T.FORUM_ID) LEFT JOIN " & strActivePrefix & "REPLY R"
strSql2 = strSql2 & " ON T.TOPIC_ID = R.TOPIC_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS M"
strSql2 = strSql2 & " ON T.T_AUTHOR = M.MEMBER_ID) LEFT JOIN " & strTablePrefix & "CATEGORY C"
strSql2 = strSql2 & " ON T.CAT_ID = C.CAT_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS MEMBERS_1"
strSql2 = strSql2 & " ON T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID"

strSql3 = " WHERE (T.T_STATUS > 1 OR R.R_STATUS > 1)"
if mlev = 3 and ModOfForums <> "" then
	strSql3 = strSql3 & " AND T.FORUM_ID IN (" & ModOfForums & ") "
end if

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
		if not (rs.EOF or rs.BOF) then
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
end if

Response.Write	"<table class=""misc"">" & vbNewline & _
				"<tr>" & vbNewline & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewline & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & " <a href=""moderate.asp"">Unmoderated Posts</a>" & vbNewLine & _
				"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine

if maxpages > 1 then
	Response.Write	"<tr>" & vbNewLine & _
					"<td colspan=""7"" class=""toppaging"">" & vbNewLine
	Call DropDownPaging(1)
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
end if

Response.Write	"<tr class=""header"">" & vbNewLine & _
				"<td>&nbsp;</td>" & vbNewLine & _
				"<td>Topic</td>" & vbNewLine & _
				"<td>Author</td>" & vbNewLine & _
				"<td>Replies</td>" & vbNewLine & _
				"<td>Read</td>" & vbNewLine & _
				"<td>Last Post</td>" & vbNewLine & _
				"<td>"
if (mlev > 0) or (lcase(strNoCookies) = "1") then
	if (mLev = 4) then
		if UnModeratedPosts > 0 then
			Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject All UnModerated Posts","hspace=""0""") & "</a>"
		else
			Response.Write("&nbsp;")
		end if
	else
		Response.Write("&nbsp;")
	end if
else
	Response.Write("&nbsp;")
end if
Response.Write	"</td>" & vbNewline & _
				"</tr>" & vbNewLine

if iTopicCount = "" then '## No Search Results
	Response.Write	"<tr>" & vbNewLine & _
					"<td colspan=""7"">No Matches Found</td>" & vbNewLine & _
					"</tr>" & vbNewLine
else 
	cCAT_STATUS = 0
	cCAT_NAME = 1
	fFORUM_ID = 4
	fF_SUBJECT = 5
	fCAT_ID = 6
	fF_PRIVATEFORUMS = 7
	fF_PASSWORD_NEW = 8
	fF_STATUS = 9
	tTOPIC_ID = 10
	tT_AUTHOR = 11
	tT_SUBJECT = 12
	tT_STATUS = 13
	tT_LAST_POST = 14
	tT_LAST_POST_AUTHOR = 15
	tT_REPLIES = 16
	tT_UREPLIES = 17
	tT_VIEW_COUNT = 18
	mMEMBER_ID = 19
	mM_NAME = 20
	tLAST_POST_AUTHOR_NAME = 21

	currForum = 0 
	currTopic = 0
	dim Cat_Status
	dim Forum_Status
	dim mdisplayed
	mdisplayed = 0
	rec = 1

	for iTopic = 0 to iTopicCount
		if (rec = strPageSize + 1) then exit for

		Cat_Status = arrTopicData(cCAT_STATUS, iTopic)
		Cat_Name = arrTopicData(cCAT_NAME, iTopic)
		Forum_ID = arrTopicData(fFORUM_ID, iTopic)
		Forum_Subject = arrTopicData(fF_SUBJECT, iTopic)
		Forum_Cat_ID = arrTopicData(fCAT_ID, iTopic)
		Forum_PrivateForums = arrTopicData(fF_PRIVATEFORUMS, iTopic)
		Forum_FPasswordNew = arrTopicData(fF_PASSWORD_NEW, iTopic)
		Forum_Status = arrTopicData(fF_STATUS, iTopic)
		Topic_ID = arrTopicData(tTOPIC_ID, iTopic)
		Topic_Author = arrTopicData(tT_AUTHOR, iTopic)
		Topic_Subject = arrTopicData(tT_SUBJECT, iTopic)
		Topic_Status = arrTopicData(tT_STATUS, iTopic)
		Topic_LastPost = arrTopicData(tT_LAST_POST, iTopic)
		Topic_LastPostAuthor = arrTopicData(tT_LAST_POST_AUTHOR, iTopic)
		Topic_Replies = arrTopicData(tT_REPLIES, iTopic)
		Topic_UReplies = arrTopicData(tT_UREPLIES, iTopic)
		Topic_ViewCount = arrTopicData(tT_VIEW_COUNT, iTopic)
		Topic_MemberID = arrTopicData(mMEMBER_ID, iTopic)
		Topic_MemberName = arrTopicData(mM_NAME, iTopic)
		Topic_LastPostAuthorName = arrTopicData(tLAST_POST_AUTHOR_NAME, iTopic)

		Dim AdminAllowed, ModerateAllowed
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
				if (instr("," & ModOfForums & "," ,"," & Forum_ID & ",") <> 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
			end if
		else
			ModerateAllowed = "N"
		end if
		if ModerateAllowed = "Y" and Topic_UReplies > 0 then
			Topic_Replies = Topic_Replies + Topic_UReplies
		end if
		if chkDisplayForum(Forum_PrivateForums,Forum_FPasswordNew,Forum_ID,MemberID) then
			if (currForum <> Forum_ID) and (currTopic <> Topic_ID) then 
				Response.Write	"<tr class=""section"">" & vbNewLine & _
								"<td colspan=""6""><a href=""default.asp?CAT_ID=" & Forum_Cat_ID & """>" & ChkString(Cat_Name,"display") & "</a>&nbsp;/&nbsp;<a href=""forum.asp?FORUM_ID=" & Forum_ID & """>" & ChkString(Forum_Subject,"display") & "</a></td>" & vbNewline
				Response.Write	"<td class=""options"">" & vbNewLine
				if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
					Call ForumAdminOptions()
  				elseif (mLev = 3) then
					Response.Write	"&nbsp;" & vbNewLine
				else
					Response.Write	"&nbsp;" & vbNewLine
				end if
				Response.Write	"</td>" & vbNewLine & _
								"</tr>" & vbNewLine

				currForum = Forum_ID
			end if 
			if currTopic <> Topic_ID then
				Response.Write	"<tr>" & vbNewline & _
								"<td class=""statusicon"">" & vbNewline
				if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then 
					' DEM --> Added if statement to display topic status properly
					if Topic_Status = 2 then
						UnApprovedFound = "Y"
						Response.Write	"<a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconFolderUnmoderated,"Topic UnModerated","hspace=""0""") & "</a>" & vbNewline
					elseif Topic_Status = 3 then
						HeldFound = "Y"
						Response.Write	"<a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconFolderHold,"Topic Held","hspace=""0""") & "</a>" & vbNewline
					else
						Response.Write	"<a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & ChkIsNew(Topic_LastPost) & "</a>" & vbNewline
					end if
				else 
 					if Cat_Status = 0 then 
 						strAltText = "Category Locked"
 					elseif Forum_Status = 0 then 
 						strAltText = "Forum Locked"
 					else
 						strAltText = "Topic Locked"
 					end if 
 					Response.Write	"<a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconFolderLocked,strAltText,"hspace=""0""") & "</a>" & vbNewline
 				end if
				Response.Write	"</td>" & vbNewline & _
								"<td><div class=""forumlink""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & ChkString(left(Topic_Subject, 50),"display") & "</a></div>" & vbNewLine
 				if strShowPaging = "1" then
 					TopicPaging()
 				end if
 				Response.Write	"</td>" & vbNewLine & _
								"<td class=""authlink"">" & profileLink(chkString(Topic_MemberName,"display"),Topic_Author) & "</td>" & vbNewLine & _
								"<td class=""counts"">" & Topic_Replies & "</td>" & vbNewLine & _
								"<td class=""counts"">" & Topic_ViewCount & "</td>" & vbNewLine
 				if IsNull(Topic_LastPostAuthor) then
 					strLastAuthor = ""
 				else
 					strLastAuthor = "<br />by:&nbsp;" & profileLink(Topic_LastPostAuthorName,Topic_LastPostAuthor)
 				end if
 				Response.Write	"<td class=""counts""><span class=""lastpost"">" & ChkDate(Topic_LastPost, "&nbsp" ,true) & _
								"</span><span class=""lastuser"">" & strLastAuthor & "</span></td>" & vbNewline
 				Response.Write	"<td class=""options"">" & vbNewLine
				if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
 					call TopicAdminOptions
 				elseif (mlev = 3) then
 					Response.Write	"&nbsp;" & vbNewLine
 				else
					Response.Write	"&nbsp;" & vbNewLine
				end if
				Response.Write	"</td>" & vbNewLine
 				Response.Write	"</tr>" & vbNewLine
 				currTopic = Topic_ID
 				rec = rec + 1 
 			end if 
 			mdisplayed = mdisplayed + 1
 		end if
 	next
 	if mdisplayed = 0 then
		Response.Write	"<tr>" & vbNewLine & _
						"<td colspan=""7"">No Matches Found</td>" & vbNewLine & _
						"</tr>" & vbNewLine
 	end if
end if 

if maxpages > 1 then
	Response.Write	"<tr>" & vbNewLine & _
					"<td colspan=""7"" class=""bottompaging"">" & vbNewLine
	Call DropDownPaging(2)
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
end if
 
Response.Write	"</table>" & vbNewLine

Response.Write	"<table class=""misc2"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td align=""left"" valign=""top"">" & vbNewLine & _
				getCurrentIcon(strIconFolderNew,"New Posts","align=""absmiddle""") & " New posts since last logon.<br />" & vbNewLine & _
				getCurrentIcon(strIconFolder,"Old Posts","align=""absmiddle""") & " Old Posts."
if lcase(strHotTopic) = "1" then Response.Write	(" (" & getCurrentIcon(strIconFolderHot,"Hot Topic","align=""absmiddle""") & "&nbsp;" & intHotTopicNum & " replies or more.)<br />" & vbNewLine)
Response.Write	getCurrentIcon(strIconFolderLocked,"Locked Topic","align=""absmiddle""") & " Locked topic.<br />" & vbNewLine
' DEM --> Start of Code added for moderation
if HeldFound = "Y" then
	Response.Write getCurrentIcon(strIconFolderHold,"Held Posts","align=""absmiddle""") & " Held Posts.<br />" & vbNewline
end if
if UnApprovedFound = "Y" then
	Response.Write getCurrentIcon(strIconFolderUnmoderated,"UnModerated Posts","align=""absmiddle""") & " UnModerated Posts.<br />" & vbNewline
end if
' DEM --> End of Code added for moderation
Response.Write	"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

WriteFooter
Response.End
 
sub ForumAdminOptions() 
	if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then 
		' DEM --> Added code to allow for moderation
		if UnModeratedPosts > 0 then
			if ModerateAllowed = "Y" and (CheckForUnModeratedPosts("FORUM", Forum_Cat_ID, Forum_ID, 0) > 0) then
				ModString = "CAT_ID=" & Forum_Cat_ID & "&FORUM_ID=" & Forum_ID
				Response.Write	"<a href=""JavaScript:openWindow('pop_moderate.asp?" & ModString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject All UnModerated Posts in this Forum","") & "</a>" & vbNewline
				UnModeratedFPosts = 1
			end if
		end if
		' DEM --> End of code added to allow for moderation
	end if 
end sub

sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then
		mxpages = int(mxpages) + 1
	end if
	if mxpages > 1 then
		Response.Write("<table class=""nb"">" & vbNewLine)
		Response.Write("<tr class=""topicpaging"">" & vbNewLine)
		Response.Write("<td>" & getCurrentIcon(strIconPosticon,"","") & "</td>" & vbNewLine)
		for counter = 1 to mxpages
			ref = "<td>" 
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
				Response.Write("<tr class=""topicpaging"">" & vbNewLine)
				Response.Write("<td>&nbsp;</td>" & vbNewLine)
			end if
		next				
		Response.Write("</tr>" & vbNewLine)
		Response.Write("</table>" & vbNewLine)
	end if
end sub

sub TopicAdminOptions()
	' DEM --> Start of Code for Full Moderation
	if UnModeratedFPosts > 0 then
	        if CheckForUnModeratedPosts("TOPIC", Forum_Cat_ID, Forum_ID, Topic_ID) > 0 then
			TopicString = "TOPIC_ID=" & Topic_ID & "&CAT_ID=" & Forum_Cat_ID & "&FORUM_ID=" & Forum_ID
                	Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject All UnModerated Posts for this Topic","hspace=""0""") & "</a>" & vbNewline
	        end if
	end if
	' DEM --> End of Code for Full Moderation 
end sub

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		If cLng(pge) > 1 then
			MinPageToShow = cLng(pge) - 1
		Else
			MinPageToShow = 1
		End If
		If cLng(pge) + strPageNumberSize > maxpages then
			MaxPageToShow = maxpages
		Else
			MaxPageToShow = cLng(pge) + strPageNumberSize
		End If
		If MaxPageToShow < maxpages then
			ShowMaxPage = True
		Else
			ShowMaxPage = False
		End If
		If MinPageToShow > 1 then
			ShowMinPage = True
		Else
			ShowMinPage = False
		End If
		
		if fnum = 1 then
			Response.Write	"<b>Page:</b>&nbsp;["
		else
			Response.Write	"<b>Total Pages (" & maxpages &"):</b>&nbsp;["
		end if
		
		If ShowMinPage then
			Response.Write	" <a href=""moderate.asp?whichpage=1&FORUM_ID=" & Forum_ID & "&sortfield=" & strtopicsortfld & _
							"&sortorder=" & strtopicsortord & _
							"""><< First</a> "
		End If
		For counter = MinPageToShow To MaxPageToShow
			if counter <> cLng(pge) then   
				Response.Write	" <a href=""moderate.asp?whichpage=" & counter & "&FORUM_ID=" & Forum_ID &"&sortfield=" & _
								strtopicsortfld & "&sortorder=" & strtopicsortord
				if strArchiveState = "1" and ArchiveView = "true" then
					Response.Write	"&ARCHIVE=" & ArchiveView
				End If	
				Response.Write	""">" & counter & "</a> "
			else
				Response.Write	" <b>[" & counter & "]</b>"
			end if
		Next
		If ShowMaxPage then
			Response.Write	" <a href=""moderate.asp?whichpage=" & maxpages & "&FORUM_ID=" & Forum_ID &"&sortfield=" & strtopicsortfld & _
							"&sortorder=" & strtopicsortord & _
							""">>> Last</a> "
		End If
		Response.Write	"]"
	end if
end sub
%>
