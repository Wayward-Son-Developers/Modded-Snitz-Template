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
%><!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#include FILE="inc_func_posting.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"--><% 

if not(strUseExtendedProfile) and Request.QueryString("verkey") = "" then 
	%><!--#INCLUDE FILE="inc_header_short.asp" --><% 
else
	%><!--#INCLUDE FILE="inc_header.asp" --><% 
end if

%><!--#INCLUDE FILE="inc_profile.asp" --><%

Dim strURLError

if Instr(1,Request.Form("refer"),"search.asp",1) > 0 then
	strRefer = "search.asp"
elseif Instr(1,Request.Form("refer"),"register.asp",1) > 0 then	
	strRefer = "default.asp"
else	
	strRefer = chkString(Request.Form("refer"),"refer")
end if
if strRefer = "" then strRefer = "default.asp"

if Request.QueryString("id") <> "" and IsNumeric(Request.QueryString("id")) = true then
	ppMember_ID = cLng(Request.QueryString("id"))
else
	ppMember_ID = 0
end if

if strAuthType = "nt" then
	if ChkAccountReg() <> "1" then 
		Response.Write	"<div class=""warning"" style=""width:50%;"">" & _
						"<p><b>Note:</b> This NT account has not been registered yet, thus the profile is not available.<br />" & vbNewLine
		if strProhibitNewMembers <> "1" then Response.Write	"If this is your account, <a href=""register.asp"">click here</a> to register."
		Response.Write	"</p></div>"
 		WriteFooter
		Response.End 
	end if
end if

'############################# E-mail Validation Mod #################################
if Request.QueryString("verkey") <> "" then
	verkey = chkString(Request.QueryString("verkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_KEY, MEMBER_ID, M_EMAIL, M_NEWEMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_KEY = '" & verkey & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		If strEmail = "1" Then
			strMessage = "<li>Your verification key did not match the one that we have in our database.<br />" & _
						"Please try changing your e-mail address again by clicking the Profile link at the top right hand corner.<br />" & _
						"If this problem persists, please <a href=""contact.asp"">contact the Administrator</a> of this forum.</li>"
		Else
			strMessage = "<li>Your verification key did not match the one that we have in our database.<br />" & _
						"Please try changing your e-mail address again by clicking the Profile link at the top right hand corner.<br />" & _
						"If this problem persists, please <a href=""mailto:" & strSender & """>contact the Administrator</a> of this forum.</li>"
		End If
		Call FailMessage(strMessage,False)
		rsKey.close
		set rsKey = nothing
		WriteFooter
		Response.End
	elseif strComp(verkey,rsKey("M_KEY")) <> 0 then
		'Error message to user
		If strEmail = "1" Then
			strMessage = "<li>Your verification key did not match the one that we have in our database.<br />" & _
						"Please try changing your e-mail address again by clicking the Profile link at the top right hand corner.<br />" & _
						"If this problem persists, please <a href=""contact.asp"">contact the Administrator</a> of this forum.</li>"
		Else
			strMessage = "<li>Your verification key did not match the one that we have in our database.<br />" & _
						"Please try changing your e-mail address again by clicking the Profile link at the top right hand corner.<br />" & _
						"If this problem persists, please <a href=""mailto:" & strSender & """>contact the Administrator</a> of this forum.</li>"
		End If
		Call FailMessage(strMessage,False)
		rsKey.close
		set rsKey = nothing
		WriteFooter
		Response.End
	elseif rsKey("M_EMAIL") = rsKey("M_NEWEMAIL") then
		If strEmail = "1" Then
			strMessage = "E-mail Already Verified!</p>" & vbNewLine & _
						"<p>Your e-mail address has already been updated in our database. " & _
						"If this problem persists, please <a href=""contact.asp"">contact the Administrator</a> of this forum.</p>" & vbNewLine & _
						"<p><a href=""default.asp"">Back To Forum</a>" & vbNewLine
		Else
			strMessage = "E-mail Already Verified!</p>" & vbNewLine & _
						"<p>Your e-mail address has already been updated in our database. " & _
						"If this problem persists, please <a href=""mailto:" & strSender & """>contact the Administrator</a> of this forum.</p>" & vbNewLine & _
						"<p><a href=""default.asp"">Back To Forum</a>" & vbNewLine
		End If
		Call PopOkMessage(strMessage,False)
		rsKey.close
		set rsKey = nothing
		WriteFooter
		Response.End
	else
		userID = rsKey("MEMBER_ID")

		'Update the user e-mail
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_EMAIL = '" & chkString(rsKey("M_NEWEMAIL"),"SQLString") & "'"
		strSql = strSql & ", M_KEY = ''"
		strSql = strSql & " WHERE MEMBER_ID = " & userID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		
		strMessage = "Your E-mail Address Has Been Updated!</p>" & vbNewLine & _
					"<p>Your new e-mail address has been successfully updated in our database.</p>" & vbNewLine & _
					"<p><a href=""default.asp"">Back To Forum</a>" & vbNewLine
		Call PopOkMessage(strMessage,False)
		
		rsKey.close
		set rsKey = nothing
		WriteFooter
		Response.End
	end if
end if
'#################################################################################

select case Request.QueryString("mode") 

	case "display" '## Display Profile
		if strDBNTUserName = "" then
			Err_Msg = "You must be logged in to view a Member's Profile"
			
			if not(strUseExtendedProfile) then
				Call FailMessage(Err_Msg,False)
				WriteFooterShort
				Response.End
			else
				Response.Write	"<table class=""misc"">" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""secondnav"">" & vbNewLine & _
								getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
								getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member's Profile</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"</table>" & vbNewLine
				Call FailMessage(Err_Msg,True)
				WriteFooter
				Response.End
			end if
		end if

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LASTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MSN"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE"
'		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_RECEIVE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK2"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DOB"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LNEWS"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_BIO"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID=" & ppMember_ID

		set rs = my_Conn.Execute(strSql)
		
		if rs.BOF or rs.EOF then
			Err_Msg = "Invalid Member ID!"
			
			if not(strUseExtendedProfile) then
				Call FailMessage(Err_Msg,False)
				WriteFooterShort
				Response.End
			else
				Response.Write	"<table class=""misc"">" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""secondnav"">" & vbNewLine & _
								getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
								getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member's Profile</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"</table>" & vbNewLine
				Call FailMessage(Err_Msg,True)
				WriteFooter
				Response.End
			end if
		else
			strMyHobbies = rs("M_HOBBIES")
			strMyQuote = rs("M_QUOTE")
			strMyLNews = rs("M_LNEWS")
			strMyBio = rs("M_BIO")

			intTotalMemberPosts = rs("M_POSTS")
			if intTotalMemberPosts > 0 then
				strMemberDays = DateDiff("d", strToDate(rs("M_DATE")), strToDate(strForumTimeAdjust))
				if strMemberDays = 0 then strMemberDays = 1
				strMemberPostsperDay = round(intTotalMemberPosts/strMemberDays,2)
				if strMemberPostsperDay = 1 then
					strPosts = " post"
				else
					strPosts = " posts"
				end if
			end if

			if strUseExtendedProfile then
				strColspan = " colspan=""2"""
				strIMURL1 = "javascript:openWindow('"
				strIMURL2 = "')"
				strICQURL = "javascript:openWindow6('"
			else
				strColspan = ""
				strIMURL1 = ""
				strIMURL2 = ""
				strICQURL = ""
			end if

			if strUseExtendedProfile then
				Response.Write	"<table class=""misc"">" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""secondnav"">" & vbNewLine & _
								getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
								getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;" & chkString(rs("M_NAME"),"display") & "'s Profile</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"</table>" & vbNewLine
			end if
			
			Response.Write	"<table class=""content"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td" & strColspan & ">User Profile</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr class=""section"">" & vbNewLine & _
							"<td" & strColspan & ">"
			if mLev = 4 then
				Response.Write	"<a href=""pop_profile.asp?mode=Modify&ID=" & rs("MEMBER_ID") & "&name=" & ChkString(rs("M_NAME"),"urlpath") & """>" & ChkString(rs("M_NAME"),"display") & "</a>" & vbNewLine
			else
				Response.Write	ChkString(rs("M_NAME"),"display") & "" & vbNewLine
			end if
			Response.Write	" - Member Since:&nbsp;" & ChkDate(rs("M_DATE"),"",false) & "</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td bgcolor=""" & strPageBGColor & """ align=""left"" valign=""top"">" & vbNewLine & _
							"<table border=""0"" width=""90%"" cellspacing=""1"" cellpadding=""0"" align=""center"">" & vbNewLine & _
							"<tr>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"<td width=""35%"" bgColor=""" & strPageBGColor & """ valign=""top"">" & vbNewLine & _
								"<table class=""admin"">" & vbNewLine
				if trim(rs("M_PHOTO_URL")) = "" or lcase(rs("M_PHOTO_URL")) = "http://" then strPicture = 0
				if strPicture = "1" then
					Response.Write	"<tr>" & vbNewLine & _
									"<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>&nbsp;My Picture&nbsp;</font></b></td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2"">"
					if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then
						Response.Write	"<a href=""" & ChkString(rs("M_PHOTO_URL"), "displayimage") & """>" & getCurrentIcon(ChkString(rs("M_PHOTO_URL"), "displayimage") & "|150|150",ChkString(rs("M_NAME"),"display"),"hspace=""2"" vspace=""2""") & "</a><br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Click image for full picture</font>"
					else
						Response.Write	getCurrentIcon(strIconPhotoNone,"No Photo Available","hspace=""2"" vspace=""2""")
					end if
					Response.Write	"</td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if ' strPicture
				Response.Write	"<tr>" & vbNewLine & _
								"<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>&nbsp;My Contact Info&nbsp;</font></b></td>" & vbNewLine & _
								"</tr>" & vbNewLine
				strContacts = 0
				if mLev > 2 or rs("M_RECEIVE_EMAIL") = "1" then
					strContacts = strContacts + 1
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" width=""10%"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>E-mail User:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_EMAIL")) <> "" then
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:openWindow('pop_mail.asp?id=" & rs("MEMBER_ID") & "')"">Click to send an E-Mail</a>&nbsp;</font></td>" & vbNewLine
					else
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No address specified...</font></td>" & vbNewLine
					end if
					Response.Write	"</tr>" & vbNewLine
				end if
				if strAIM = "1" and Trim(rs("M_AIM")) <> "" then 
					strContacts = strContacts + 1
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>AIM:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconAIM,"","align=""absmiddle""") & "&nbsp;<a href=""" & strIMURL1 & "pop_messengers.asp?mode=AIM&ID=" & rs("MEMBER_ID") & strIMURL2 & """>" & ChkString(rs("M_AIM"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if 
				if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then 
					strContacts = strContacts + 1
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>ICQ:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & ChkString(rs("M_ICQ"), "urlpath") & "&img=5|18|18","","align=""absmiddle""") & "&nbsp;<a href=""" & strICQURL & "pop_messengers.asp?mode=ICQ&ID=" & rs("MEMBER_ID") & strIMURL2 & """>" & ChkString(rs("M_ICQ"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strMSN = "1" and Trim(rs("M_MSN")) <> "" then 
					strContacts = strContacts + 1
					parts = split(rs("M_MSN"),"@")
					strtag1 = parts(0)
					partss = split(parts(1),".")
					strtag2 = partss(0)
					strtag3 = ""
					for xmsn = 1 to ubound(partss)
						if strtag3 <> "" then strtag3 = strtag3 & "."
						strtag3 = strtag3 & partss(xmsn)
					next

					Response.Write	"<script language=""javascript"" type=""text/javascript"">" & vbNewLine & _
									"    function MSNjs() {" & vbNewLine & _
									"        var tag1 = '" & strtag1 & "';" & vbNewLine & _
									"        var tag2 = '" & strtag2 & "';" & vbNewLine & _
									"        var tag3 = '" & strtag3 & "';" & vbNewLine & _
									"        document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & vbNewLine & _
									"</script>" & vbNewLine

					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>MSN:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconMSNM,"","align=""absmiddle""") & "&nbsp;<script language=""javascript"" type=""text/javascript"">MSNjs()</script>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then 
					strContacts = strContacts + 1
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>YAHOO IM:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(rs("M_YAHOO"), "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon("http://opi.yahoo.com/online?u=" & ChkString(rs("M_YAHOO"), "urlpath") & "&m=g&t=2|125|25","","") & "</a>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strContacts = 0 then
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No info specified...</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strRecentTopics = "1" then
					strStartDate = DateToStr(dateadd("d", -30, strForumTimeAdjust))

					'## Forum_SQL - Find all records for the member
					strsql = "SELECT F.FORUM_ID"
					strSql = strSql & ", T.TOPIC_ID"
					strSql = strSql & ", T.T_SUBJECT"
					strSql = strSql & ", T.T_STATUS"
					strSql = strSql & ", T.T_LAST_POST"
					strSql = strSql & ", T.T_REPLIES "
					strSql = strSql & " FROM ((" & strTablePrefix & "FORUM F LEFT JOIN " & strTablePrefix & "TOPICS T"
					strSql = strSql & " ON F.FORUM_ID = T.FORUM_ID) LEFT JOIN " & strTablePrefix & "REPLY R"
					strSql = strSql & " ON T.TOPIC_ID = R.TOPIC_ID) "
					strSql = strSql & " WHERE (T_DATE > '" & strStartDate & "') "
					strSql = strSql & " AND (T.T_AUTHOR = " & ppMember_ID
					strSql = strSql & " OR R.R_AUTHOR = " & ppMember_ID & ")"
					strSql = strSql & " AND (T_STATUS < 2 OR R_STATUS < 2)"
					strSql = strSql & " AND F.F_TYPE = 0"
					strSql = strSql & " ORDER BY T.T_LAST_POST DESC, T.TOPIC_ID DESC"

					set rs2 = my_Conn.Execute(strsql)					

					Response.Write	"<tr>" & vbNewLine & _
									"<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Recent Topics</font></b></td>" & vbNewLine & _
									"</tr>" & vbNewLine
					if rs2.EOF or rs2.BOF then
						Response.Write	"<tr>" & vbNewLine & _
										"<td bgcolor=""" & strPopUpTableColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<br />&nbsp;<b>No Matches Found...<br />&nbsp;</b></font></td>" & vbNewLine & _
										"</tr>" & vbNewLine
					else 
						currTopic = 0
						TopicCount = 0
						Response.Write	"<tr>" & vbNewLine & _
										"<td bgColor=""" & strPopUpTableColor & """ valign=""top"" colspan=""2"">" & vbNewLine & _
										"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine
						do until rs2.EOF or (TopicCount = 10)
							if chkForumAccess(rs2("FORUM_ID"),MemberID,false) then 
								if currTopic <> rs2("TOPIC_ID") then 
									Response.Write	"<tr>" & vbNewLine & _
													"<td bgcolor=""" & strPopUpTableColor & """ width=""5%"">" & vbNewLine & _
													"<a href=""topic.asp?TOPIC_ID=" & rs2("TOPIC_ID") & """>"
									if rs2("T_STATUS") <> 0 then
										if strHotTopic = "1" then
											if rs2("T_LAST_POST") > Session(strCookieURL & "last_here_date") then
												if rs2("T_REPLIES") >= intHotTopicNum then
													Response.Write	getCurrentIcon(strIconFolderNewHot,"Hot Topic","align=""absmiddle""") & "</a></td>" & vbNewLine
												else
													Response.Write	getCurrentIcon(strIconFolderNew,"New Topic","align=""absmiddle""") & "</a></td>" & vbNewLine
												end if
											else
												if rs2("T_REPLIES") >= intHotTopicNum then
													Response.Write	getCurrentIcon(strIconFolderHot,"Hot Topic","align=""absmiddle""") & "</a></td>" & vbNewLine
												else
													Response.Write	getCurrentIcon(strIconFolder,"","align=""absmiddle""") & "</a></td>" & vbNewLine
												end if
											end if
										else
											if rs2("T_LAST_POST") > Session(strCookieURL & "last_here_date") then
												Response.Write	getCurrentIcon(strIconFolderNew,"New Topic","align=""absmiddle""") & "</a></td>" & vbNewLine
											else
												Response.Write	getCurrentIcon(strIconFolder,"","align=""absmiddle""") & "</a></td>" & vbNewLine
											end if
										end if
									else 
										if rs2("T_LAST_POST") > Session(strCookieURL & "last_here_date") then 
											Response.Write	getCurrentIcon(strIconFolderNewLocked,"Topic Locked","align=""absmiddle""") & "</a></td>" & vbNewLine
										else
											Response.Write	getCurrentIcon(strIconFolderLocked,"Topic Locked","align=""absmiddle""") & "</a></td>" & vbNewLine
										end if
									end if 
									Response.Write	"<td bgcolor=""" & strPopUpTableColor & """ align=""left"" valign=""middle"" width=""95%""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>&nbsp;<a href=""topic.asp?TOPIC_ID=" & rs2("TOPIC_ID") & """>" & ChkString(rs2("T_SUBJECT"),"display") & "</a>&nbsp;</font></td>" & vbNewLine & _
													"</tr>" & vbNewLine
									TopicCount = TopicCount + 1
								end if 
								currTopic = rs2("TOPIC_ID")
							end if
							rs2.MoveNext 
						loop 
						Response.Write	"</table>" & vbNewLine & _
										"</td>" & vbNewLine & _
										"</tr>" & vbNewLine
					end if 
					rs2.close
					set rs2 = nothing

				elseif (strHomepage + strFavLinks) > 0 and (strRecentTopics = "0") then  

					Response.Write	"<tr>" & vbNewLine & _
									"<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2"">" & vbNewLine & _
									"<b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Links&nbsp;</font></b></td>" & vbNewLine
					if strHomepage = "1" then
						Response.Write	"<tr>" & vbNewLine & _
										"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Homepage:&nbsp;</font></b></td>" & vbNewLine
						if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then
							Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & rs("M_HOMEPAGE") & """ target=""_blank"">" & rs("M_HOMEPAGE") & "</a>&nbsp;</font></td>" & vbNewLine
						else
							Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No homepage specified...</font></td>" & vbNewLine
						end if
						Response.Write	"</tr>" & vbNewLine
					end if
					if strFavLinks = "1" then 
						Response.Write	"<tr>" & vbNewLine & _
										"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Cool Links:&nbsp;</font></b></td>" & vbNewLine
						if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then
							Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & rs("M_LINK1") & """ target=""_blank"">" & rs("M_LINK1") & "</a>&nbsp;</font></td>" & vbNewLine
							if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then
								Response.Write	"</tr>" & vbNewLine & _
												"<tr>" & vbNewLine & _
												"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></b></td>" & vbNewLine & _
												"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & rs("M_LINK2") & """ target=""_blank"">" & rs("M_LINK2") & "</a>&nbsp;</font></td>" & vbNewLine
							end if
						else
							Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No link specified...</font></td>" & vbNewLine
						end if 
						Response.Write	"</tr>" & vbNewLine
					end if
				end if ' strRecentTopics
				Response.Write	"</table>" & vbNewLine & _
								"</td>" & vbNewLine & _
								"<td valign=""top"" width=""3%"" bgColor=""" & strPageBGColor & """>&nbsp;</td>" & vbNewLine
			end if ' UseExtendedMemberProfile
			Response.Write	"<td bgColor=""" & strPageBGColor & """ valign=""top"">" & vbNewLine & _
							"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""3"" valign=""top"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td valign=""top"" align=""center"" colspan=""2"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Basics</font></b></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%"" valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:&nbsp;</font></b></td>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_NAME"),"display") & "&nbsp;</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine
			if strAuthType = "nt" then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Account:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_USERNAME"),"display") & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if strFullName = "1" and (Trim(rs("M_FIRSTNAME")) <> "" or Trim(rs("M_LASTNAME")) <> "" ) then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Real Name:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_FIRSTNAME"), "display") & "&nbsp;" & ChkString(rs("M_LASTNAME"), "display") & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if (strCity = "1" and Trim(rs("M_CITY")) <> "") or (strCountry = "1" and Trim(rs("M_COUNTRY")) <> "") or (strState = "1" and Trim(rs("M_STATE")) <> "") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Location:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
				myCity = ChkString(rs("M_CITY"),"display")
				myState = ChkString(rs("M_STATE"),"display")
				myCountry = ChkString(rs("M_COUNTRY"),"display")
				myLocation = ""

				if myCity <> "" and myCity <> " " then
					myLocation = myCity
				end if

				if myLocation <> "" then
					if myState <> "" and myState <> " " then
						myLocation = myLocation & ",&nbsp;" & myState
					end if
				else
					if myState <> "" and myState <> " " then
						myLocation = myState
					end if
				end if

				if myLocation <> "" then
					if myCountry <> "" and myCountry <> " " then
						myLocation = myLocation & "<br />" & myCountry
					end if
				else
					if myCountry <> "" and myCountry <> " " then
						myLocation = myCountry
					end if
				end if
				Response.Write	myLocation
				Response.Write	"</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if (strAge = "1" and Trim(rs("M_AGE")) <> "") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Age:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_AGE"), "display") & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			strDOB = rs("M_DOB")
			if (strAgeDOB = "1" and Trim(strDOB) <> "") then
			strDOB = DOBToDate(strDOB)
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Age:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & DisplayUsersAge(strDOB) & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if (strMarStatus = "1" and Trim(rs("M_MARSTATUS")) <> "") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Marital Status:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_MARSTATUS"), "display") & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if (strSex = "1" and Trim(rs("M_SEX")) <> "") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Gender:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_SEX"), "display") & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if (strOccupation = "1" and Trim(rs("M_OCCUPATION")) <> "") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Occupation:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_OCCUPATION"), "display") & "</font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if intTotalMemberPosts > 0 then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Total Posts:&nbsp;</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(intTotalMemberPosts, "display") & "<br /><font size=""" & strFooterFontSize & """>[" & strMemberPostsperDay & strPosts & " per day]<br /><a href=""search.asp?mode=DoIt&MEMBER_ID=" & rs("MEMBER_ID") & """>Find all non-archived posts by " & chkString(rs("M_NAME"),"display") & "</a></font></font></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			if not(strUseExtendedProfile) then
				if rs("M_RECEIVE_EMAIL") = "1" then
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" width=""10%"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>E-mail User:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_EMAIL")) <> "" then
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""pop_mail.asp?id=" & rs("MEMBER_ID") & """>Click to send an E-Mail</a>&nbsp;</font></td>" & vbNewLine
					else
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No address specified...</font></td>" & vbNewLine
					end if
					Response.Write	"</tr>" & vbNewLine
				end if
				if strAIM = "1" and Trim(rs("M_AIM")) <> "" then
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>AIM:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconAIM,"","align=""absmiddle""") & "&nbsp;<a href=""pop_messengers.asp?mode=AIM&ID=" & rs("MEMBER_ID") & """>" & ChkString(rs("M_AIM"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if 
				if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then 
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>ICQ:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & ChkString(rs("M_ICQ"), "urlpath") & "&img=5|18|18","","align=""absmiddle""") & "&nbsp;<a href=""pop_messengers.asp?mode=ICQ&ID=" & rs("MEMBER_ID") & """>" & ChkString(rs("M_ICQ"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strMSN = "1" and Trim(rs("M_MSN")) <> "" then
					parts = split(rs("M_MSN"),"@")
					strtag1 = parts(0)
					partss = split(parts(1),".")
					strtag2 = partss(0)
					strtag3 = partss(1)

					Response.Write	"<script language=""javascript"" type=""text/javascript"">" & vbNewLine & _
									"    function MSNjs() {" & vbNewLine & _
									"        var tag1 = '" & strtag1 & "';" & vbNewLine & _
									"        var tag2 = '" & strtag2 & "';" & vbNewLine & _
									"        var tag3 = '" & strtag3 & "';" & vbNewLine & _
									"        document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & vbNewLine & _
									"</script>" & vbNewLine

					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>MSN:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconMSNM,"","align=""absmiddle""") & "&nbsp;<script language=""javascript"" type=""text/javascript"">MSNjs()</script>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if 
				if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then 
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>YAHOO IM:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(rs("M_YAHOO"), "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon("http://opi.yahoo.com/online?u=" & ChkString(rs("M_YAHOO"), "urlpath") & "&m=g&t=2|125|25","","") & "</a>&nbsp;</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
			end if
			if IsNull(strMyBio) or trim(strMyBio) = "" then strBio = 0
			if IsNull(strMyHobbies) or trim(strMyHobbies) = "" then strHobbies = 0
			if IsNull(strMyLNews) or trim(strMyLNews) = "" then strLNews = 0
			if IsNull(strMyQuote) or trim(strMyQuote) = "" then strQuote = 0
			if (strBio + strHobbies + strLNews + strQuote) > 0 then
				Response.Write	"<tr>" & vbNewLine & _
								"<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>More About Me</font></b></td>" & vbNewLine & _
								"</tr>" & vbNewLine
				if strBio = "1" then
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Bio:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyBio) or trim(strMyBio) = "" then Response.Write("-") else Response.Write(formatStr(strMyBio))
					Response.Write	"</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strHobbies = "1" then  
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Hobbies:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyHobbies) or trim(strMyHobbies) = "" then Response.Write("-") else Response.Write(formatStr(strMyHobbies))
					Response.Write	"</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strLNews = "1" then  
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Latest News:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyLNews) or trim(strMyLNews) = "" then Response.Write("-") else Response.Write(formatStr(strMyLNews))
					Response.Write	"</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strQuote = "1" then  
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgcolor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Favorite Quote:&nbsp;</font></b></td>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyQuote) or Trim(strMyQuote) = "" then Response.Write("-") else Response.Write(formatStr(strMyQuote))
					Response.Write	"</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
			end if
			if (strHomepage + strFavLinks) > 0 and not(strRecentTopics = "0" and strUseExtendedProfile) then  
				if strUseExtendedProfile then	
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgcolor=""" & strCategoryCellColor & """ align=""center"" colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Links&nbsp;</font></b></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				end if
				if strHomepage = "1" then 
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Homepage:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & ChkString(rs("M_HOMEPAGE"), "display") & """ target=""_blank"">" & ChkString(rs("M_HOMEPAGE"), "display") & "</a>&nbsp;</font></td>" & vbNewLine
					else
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No homepage specified...</font></td>" & vbNewLine
					end if
					Response.Write	"</tr>" & vbNewLine
				end if
				if strFavLinks = "1" then
					Response.Write	"<tr>" & vbNewLine & _
									"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Cool Links:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then 
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & ChkString(rs("M_LINK1"), "display") & """ target=""_blank"">" & ChkString(rs("M_LINK1"), "display") & "</a>&nbsp;</font></td>" & vbNewLine
						if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then
							Response.Write	"</tr>" & vbNewLine & _
											"<tr>" & vbNewLine & _
											"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></b></td>" & vbNewLine & _
											"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & ChkString(rs("M_LINK2"), "display") & """ target=""_blank"">" & ChkString(rs("M_LINK2"), "display") & "</a>&nbsp;</font></td>" & vbNewLine
						end if
					else
						Response.Write	"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No link specified...</font></td>" & vbNewLine
					end if 
					Response.Write	"</tr>" & vbNewLine
				end if
			end if
			if mlev = 4 or (mlev = 3 and CInt(strUGModForums) > 0 ) or (mlev > 0 and CInt(strUGMemView) = 2) or (ppMember_ID = MemberID and CInt(strUGMemView) > 0) then
				strSql = "SELECT U.USERGROUP_NAME, U.MEM_HIDE FROM " & strTablePrefix & "USERGROUPS U, " &_
					strTablePrefix & "USERGROUP_USERS UM " &_
					"WHERE U.USERGROUP_ID = UM.USERGROUP_ID " &_
					"AND UM.MEMBER_TYPE = 1 " &_
					"AND UM.MEMBER_ID = " & ppMember_ID & " " &_
					"ORDER BY U.USERGROUP_NAME"
				set rsPPGroups = my_Conn.execute(strSql)
				arPPGroups = Null
				if not rsPPGroups.bof and not rsPPGroups.eof then arPPGroups = rsPPGroups.GetRows
				rsPPGroups.close
				set rsPPGroups = Nothing
				strPPGroups = ""
				if not IsNull(arPPGroups) then
					for iPPGcount = LBound(arPPGroups,2) to UBound(arPPGroups,2)
						if (arPPGroups(1,iPPGcount) = 0 or mlev = 4) then
							if strPPGroups <> "" then strPPGroups = strPPGroups & ", "
							strPPGroups = strPPGroups & arPPGroups(0,iPPGcount)
						end if
					next
				end if
				Response.Write	"<tr>" & vbNewLine & _
								"<td valign=""top"" align=""center"" colspan=""2"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>UserGroups</font></b></td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td colspan=""2"" bgColor=""" & strPopUpTableColor & """><table><tr>"
				if CInt(strUGView) > 0 then response.write "<td><a href=""usergroups.asp?mode=ViewGroups"" alt=""View UserGroups"">" & getCurrentIcon(strIconGroup,"View UserGroups","hspace=""0""") & "</a></td>"
				Response.Write    "<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
				if strPPGroups = "" then
					Response.Write	"no usergroups found"
				else
					Response.Write    ChkString(strPPGroups,"display")
				end if
				Response.Write	"</font></td>" & vbNewline &_
								"</tr></table></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
			Response.Write	"</table>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"</table>" & vbNewLine & _
								"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td>" & _
								"<br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Back to previous page</a></font></p><br />" & vbNewLine
			else
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPageBGColor & """ align=""center"" nowrap>" & vbNewLine
			end if
		end if
	
	case "Edit" 
		if strUseExtendedProfile then
			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
							getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
							getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Edit Your Profile</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		end if
		Response.Write	"<center>" & vbNewLine & _
						"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>User Profile</font></p>" & vbNewLine & _
						"<p align=""center""><form action=""pop_profile.asp?mode=goEdit"" name=""goEdit"" method=""post"">" & vbNewLine & _
						"<input name=""Refer"" type=""hidden"" value=""" & strReferer & """>" & vbNewLine & _
						"<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>It is up to you to keep your profile up to date.<br />" & vbNewLine
		if strAuthType = "nt" then
			Response.Write	"Your NT account is shown. Click Submit to carry on.<br /><br />" & vbNewLine
		else 
			if strAuthType = "db" then 
				Response.Write	"Please Fill the Form in with your details.<br /><br />" & vbNewLine
			end if
		end if 
		if strProhibitNewMembers <> "1" and MemberID < 0 then
			Response.Write	"If you have not registered then <a href=""register.asp"">do so here</a>.</font></center></p>" & vbNewLine
		else
			Response.Write	"</font></center></p>" & vbNewLine
		end if

		Response.Write	"<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
						"<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine
		if strAuthType = "nt" then
			Response.Write	"<tr>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Account:</font></b></td>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>" & Session(strCookieURL & "userid") & "</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		else
			if strAuthType = "db" then
				Response.Write	"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><input name=""Name"" size=""25"" value=""" & chkString(strDBNTUserName,"display") & """ style=""width:150px;""></td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:</font></b></td>" & vbNewLine & _
								"<td bgColor=""" & strPopUpTableColor & """><input name=""Password"" type=""Password"" size=""25"" value="""" style=""width:150px;""></td>" & vbNewLine & _
								"</tr>" & vbNewLine
				if strDBNTUserName <> "" then
					Response.Write	"<script language=""JavaScript"" type=""text/javascript"">document.goEdit.Password.focus();</script>" & vbNewLine
				else
					Response.Write	"<script language=""JavaScript"" type=""text/javascript"">document.goEdit.Name.focus();</script>" & vbNewLine
				end if
			end if
		end if 
		Response.Write	"<tr>" & vbNewLine & _
						"<td bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2""><input type=""submit"" value=""Submit""></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</form>" & vbNewLine
	
	case "goEdit"
		if strAuthType = "db" then
			if strDBNTUserName = "" then 
				strDBNTUserName = Request.Form("Name")
			end if
		end if

		strEncodedPassword = sha256("" & Request.Form("Password"))

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LASTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MSN"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE"
'		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_RECEIVE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK2"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DOB"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_VIEW_SIG"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG_DEFAULT"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LNEWS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_BIO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
		strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(strDBNTUserName, "SQLString") & "' "
		if strAuthType = "db" then
			strSql = strSql & " AND   M_PASSWORD = '" & ChkString(strEncodedPassword,"SQLString") & "'"
		end if

		set rs = my_Conn.Execute(strSql)

		if strUseExtendedProfile then
			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
							"" & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
							"" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Edit Your Profile</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		end if
		if rs.BOF or rs.EOF or not(ChkQuoteOk(strDBNTUserName)) or not(ChkQuoteOk(strEncodedPassword)) then 
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>Invalid UserName or Password</font></p>" & vbNewLine & _
							"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back To Retry</a></font></p>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
			end if 
		else
			'## Display Edit Profile Page
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Edit User Profile</font></p>" & vbNewLine & _
							"<p align=""center""><form action=""pop_profile.asp?mode=EditIt"" method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
							"<input name=""Refer"" type=""hidden"" value=""" & chkString(Request.Form("Refer"),"refer") & """>" & vbNewLine
			Call DisplayProfileForm
			Response.Write	"</form></p>" & vbNewLine
		end if 
	case "Modify"
		if strUseExtendedProfile then
			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
							"" & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
							"" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Modify " & GetMemberName(ppMember_ID) & "'s Profile</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		end if
		Response.Write	"<center>" & vbNewLine & _
						"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Modify Member</font></p>" & vbNewLine
		if ppMember_ID = cLng(intAdminMemberID) and cLng(MemberID) <> cLng(intAdminMemberID) then
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b><font color=""" & strHiLiteFontColor & """>NOTE:</font></b> The <b>Forum Admin</b> account can only be modified by the Forum Admin.</font></p>" & vbNewLine & _
							"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Back to Forum</a></font></p>" & vbNewLine
		else
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b><font color=""" & strHiLiteFontColor & """>NOTE:</font></b> Only Administrators can Modify a Member.</font></p>" & vbNewLine & _
							"<form action=""pop_profile.asp?mode=goModify"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
							"<input type=""hidden"" name=""Method_Type"" value=""" & Request.QueryString("mode") & """>" & vbNewLine & _
							"<input type=""hidden"" name=""MEMBER_ID"" value=""" & ppMember_ID & """>" & vbNewLine & _
							"<input type=""hidden"" name=""Refer"" value=""" & strReferer & """>" & vbNewLine & _
							"<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
							"<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine
		if strAuthType="db" then
			Response.Write	"<tr>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:</font></b></td>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""User"" value=""" & chkString(strDBNTUserName,"display") & """ size=""20"" style=""width:150px;""></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:</FONT></b></td>" & vbNewLine & _
							"<td bgColor=""" & strPopUpTableColor & """><input type=""Password"" name=""Pass"" value="""" size=""20"" style=""width:150px;""></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		elseif strAuthType="nt" then
			Response.Write	"<tr>" & vbNewLine & _
							"<td bgcolor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>NT Account:</font></b></td>" & vbNewLine & _
							"<td bgcolor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & Session(strCookieURL & "userid") & "</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		end if
		Response.Write	"<tr>" & vbNewLine & _
						"<td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""Submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</form>" & vbNewLine
		end if
	
	case "goModify"
		if strNoCookies = "1" and strAuthType = "db" then
			if strDBNTUserName = "" then 
				strDBNTUserName = chkString(Request.Form("User"),"SQLString")
			end if
		end if

		strEncodedPassword = sha256("" & Request.Form("Pass"))
		mLev = cLng(chkUser(strDBNTUserName, strEncodedPassword,-1)) 
						
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				'## Forum_SQL
				strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_EMAIL"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_IP"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LAST_IP"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LASTNAME"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MSN"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE"
'				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_RECEIVE_EMAIL"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK2"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DOB"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_VIEW_SIG"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG_DEFAULT"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LNEWS"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE"
				strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_BIO"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
				strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ALLOWEMAIL"
				strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
				strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))
			
				set rs = my_Conn.Execute(strSql)

				if rs("M_LEVEL") = 3 then
					if cLng(MemberID) = cLng(rs("MEMBER_ID")) OR cLng(MemberID) = cLng(intAdminMemberID) then
						'Do Nothing
					else
						rs.close
						set rs = nothing
						Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Modify an Administrator</b></font><br />" & vbNewLine & _
										"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
						if strUseExtendedProfile then
							Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
							WriteFooter
							Response.End
						else
							WriteFooterShort
							Response.End
						end if
					end if
				end if
				if strUseExtendedProfile then
					Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
									"" & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
									"" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Modify " & chkString(rs("M_NAME"),"display") & "'s Profile</font></td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"</table>" & vbNewLine
				end if
				'## Display Edit Profile Page
				Response.Write	"<center>" & vbNewLine & _
								"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Modify User Profile</font></p>" & vbNewLine & _
								"<p align=""center""><form action=""pop_profile.asp?mode=ModifyIt&id=" & Request.Form("MEMBER_ID") & """ method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
								"</center>" & vbNewLine & _
								"<input type=""hidden"" name=""User"" value=""" & strDBNTUserName & """>" & vbNewLine & _
								"<input type=""hidden"" name=""Pass"" value=""" & strEncodedPassword & """>" & vbNewLine & _
								"<input type=""hidden"" name=""Refer"" value=""" & Request.Form("Refer") & """>" & vbNewLine
				Call DisplayProfileForm
				Response.Write	"</form></p>" & vbNewLine
			else
				Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Modify a Member</b></font><br />" & vbNewLine & _
								"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
				if strUseExtendedProfile then
					Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
				end if 
			end if 
		else 
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Modify a Member</b></font><br />" & vbNewLine & _
							"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
			end if 
		end if 
	
	case "EditIt"
		if strSignatures = "1" then
			intSigDefault = Request.Form("fSigDefault")
			Session(strCookieURL & "intSigDefault" & MemberID) = intSigDefault
			Session(strCookieURL & "intSigDefault" & MemberID) = intSigDefault
		end if
		if strUseExtendedProfile then
			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
							"" & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
							"" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Edit Your Profile</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		end if

		Err_Msg = ""
		if trim(Request.Form("Name")) = "" then 
			Err_Msg = Err_Msg & "<li>You must choose a UserName</li>"
		end if
		if strMSN = "1" and strReqMSN = "1" then
			if trim(Request.Form("MSN")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Valid MSN Name</li>"
			end if
		end if
		if strAIM = "1" and strReqAIM = "1" then
			if trim(Request.Form("AIM")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Valid AIM Name</li>"
			end if
		end if
		if strICQ = "1" and strReqICQ = "1" then
			if trim(Request.Form("ICQ")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Valid ICQ Name</li>"
			end if
		end if
		if strYAHOO = "1" and strReqYAHOO = "1" then
			if trim(Request.Form("YAHOO")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Valid Yahoo! Name</li>"
			end if
		end if
		if strFullName = "1" and strReqFullName = "1" then
			if trim(Request.Form("FirstName")) = "" then
				Err_Msg = Err_Msg & "<li>You must provide your First Name</li>"
			end if
			if trim(Request.Form("LastName")) = "" then
				Err_Msg = Err_Msg & "<li>You must provide your Last Name</li>"
			end if
		end if
		if strPicture = "1" and strReqPicture = "1" then
			if trim(Request.Form("Photo_URL")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Picture</li>"
			end if
		end if
		if strSex = "1" and strReqSex = "1" then
			if trim(Request.Form("Sex")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Gender</li>"
			end if
		end if
		if strCity = "1" and strReqCity = "1" then
			if trim(Request.Form("City")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your City</li>"
			end if
		end if
		if strState = "1" and strReqState = "1" then
			if trim(Request.Form("State")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your State</li>"
			end if
		end if
		if strAge = "1" and strReqAge = "1" then
			if trim(Request.Form("Age")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Age</li>"
			end if
		end if
		if strHomepage = "1" and strReqHomepage = "1" then
			if trim(Request.Form("Homepage")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Homepage</li>"
			end if
		end if
		if strCountry = "1" and strReqCountry = "1" then
			if trim(Request.Form("Country")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Country</li>"
			end if
		end if
		if strOccupation = "1" and strReqOccupation = "1" then
			if trim(Request.Form("Occupation")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Occupation</li>"
			end if
		end if
		if strBio = "1" and strReqBio = "1" then
			if trim(Request.Form("Bio")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Bio</li>"
			end if
		end if
		if strHobbies = "1" and strReqHobbies = "1" then
			if trim(Request.Form("Hobbies")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Hobbies</li>"
			end if
		end if
		if strLNEWS = "1" and strReqLNEWS = "1" then
			if trim(Request.Form("LNEWS")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Latest News</li>"
			end if
		end if
		if strQuote = "1" and strReqQuote = "1" then
			if trim(Request.Form("Quote")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide A Quote</li>"
			end if
		end if
		if strMarStatus = "1" and strReqMarStatus = "1" then
			if trim(Request.Form("MarStatus")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Marital Status</li>"
			end if
		end if
		if strFavLinks = "1" and strReqFavLinks = "1" then
			if trim(Request.Form("Link1")) = "" and trim(Request.Form("Link2")) = "" then
				Err_Msg = Err_Msg & "<li>You Must Provide at Least One Cool Link</li>"
			end if
		end if
		if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) then
			Err_Msg = Err_Msg & "<li> &gt; and &lt; are not allowed in the UserName, Please Choose Another</li>"
		end if
		if strAuthType = "db" then
			if trim(Request.Form("Password")) <> "" then 
				if Len(Request.Form("Password")) > 25 then 
					Err_Msg = Err_Msg & "<li>Your Password can not be greater than 25 characters</li>" 
				end if
				if Request.Form("Password") <> Request.Form("Password2") then 
					Err_Msg = Err_Msg & "<li>Your Passwords didn't match.</li>"
				end if
			end if
		end if
		If Request.Form("Email") = "" Then
			Err_Msg = Err_Msg & "<li>You Must give an e-mail address</li>"
		Else
			'Comment out down to the next comment to let it take me@example.com and/or .ex as well
			'strsql = "SELECT SPAM_SERVER FROM " & strTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '" & chkString(Request.Form("Email"),"sqlstring") & "'"
			'set rsSpam = my_Conn.Execute (strsql)
			
			'If Not rsSpam.EOF Then
			'	Err_Msg = Err_Msg & "<li>You cannot register with '" & chkString(Request.Form("Email"),"sqlstring") & "'.</li>"
			'End If 
			
			'Dim strMailTLD : strMailTLD = LCase(Mid(Request.Form("Email"),InStrRev(Request.Form("Email"),".")))
			
			'strsql = "SELECT SPAM_SERVER FROM " & strTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '" & strMailTLD & "'"
			'set rsSpam = my_Conn.Execute (strsql)
			
			'If Not rsSpam.EOF Then
			'	Err_Msg = Err_Msg & "<li>You cannot register with a '" & strMailTLD & "' email address.</li>"
			'End If 
			'Comment out up to the previous comment to let it take me@example.com and/or .ex as well
			
			strMailDomain = LCase(Mid(Request.Form("Email"),InStrRev(Request.Form("Email"),"@")))
			
			strsql = "SELECT SPAM_SERVER FROM " & strTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '" & strMailDomain & "'"
			set rsSpam = my_Conn.Execute (strsql)
			
			If Not rsSpam.EOF Then
				Err_Msg = Err_Msg & "<li>You cannot register with an '" & strMailDomain & "' email address.</li>"
			End If 
			
			rsSpam.close
			Set rsSpam = Nothing
		End If
		if EmailField(Request.Form("Email")) = 0 then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address</li>"
		end if
		if strMSN = "1" and trim(Request.Form("MSN")) <> "" then
			if EmailField(Request.Form("MSN")) = 0 then 
				Err_Msg = Err_Msg & "<li>You Must enter a valid MSN Messenger Username</li>"
			end if
		end if

		if strUniqueEmail = "1" then
			if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) then
				'## Forum_SQL
				strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " WHERE M_EMAIL = '" & Trim(ChkString(Request.Form("Email"), "SQLString")) &"'"

				set rs = my_Conn.Execute(TopSQL(strSql,1))

				if rs.BOF and rs.EOF then 
					'## Do Nothing - proceed
				else 
					Err_Msg = Err_Msg & "<li>E-mail Address already in use, Please Choose Another</li>"
				end if
				set rs = nothing

				if strEmail = "1" and strEmailVal = "1" then
					'## Forum_SQL
					strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
					strSql = strSql & " WHERE M_EMAIL = '" & Trim(ChkString(Request.Form("Email"),"SQLString")) &"'"

					set rs = my_Conn.Execute(TopSQL(strSql,1))

					if rs.BOF and rs.EOF then 
						'## Do Nothing
					else
						Err_Msg = Err_Msg & "<li>E-mail Address already in use, Please Choose Another</li>"
					end if
					set rs = nothing

					'## Forum_SQL
					strSql = "SELECT M_NEWEMAIL FROM " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " WHERE M_NEWEMAIL = '" & Trim(ChkString(Request.Form("Email"),"SQLString")) &"'"

					set rs = my_Conn.Execute(TopSQL(strSql,1))

					if rs.BOF and rs.EOF then 
						'## Do Nothing
					else
						Err_Msg = Err_Msg & "<li>E-mail Address already in use, Please Choose Another</li>"
					end if
					set rs = nothing
				end if
				if lcase(strEmail) = "1" and Err_Msg = "" and strEmailVal = "1" then
					verKey= GetKey("sendemail")
				end if
			end if
		else
			if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) and lcase(strEmail) = "1" and strEmailVal = "1" then
				verKey = GetKey("sendemail")
			end if
		end if
		if not IsValidURL(trim(Request.Form("Homepage"))) then
			Err_Msg = Err_Msg & "<li>Homepage URL: Invalid URL" & strURLError & "</li>"
		end if
		if not IsValidURL(trim(Request.Form("LINK1"))) then
			Err_Msg = Err_Msg & "<li>Cool Links URL: Invalid URL" & strURLError & "</li>"
		end if
		if not IsValidURL(trim(Request.Form("LINK2"))) then
			Err_Msg = Err_Msg & "<li>Cool Links URL: Invalid URL" & strURLError & "</li>"
		end if
		if not IsValidURL(trim(Request.Form("Photo_URL"))) then
			Err_Msg = Err_Msg & "<li>Photo URL: Invalid URL" & strURLError & "</li>"
		end if
		strMAge = ""
		if strAge = "1" then
			strMAge = ChkString(trim(Request.Form("Age")), "SQLString")
		end if
		if strAgeDOB = "1" then
			strMDOB = ChkString(Request.Form("year"), "SQLString") & ChkString(Request.Form("month"), "SQLString") & ChkString(Request.Form("day"), "SQLString")
			if len(strMDOB) <> 8 then
				strMDOB = ""
			else
				if not IsValidBirthDate() then
					Err_Msg = Err_Msg & "<li>Date of Birth: Invalid Date</li>"
				else
					strMAge = DisplayUsersAge(DOBToDate(strMDOB))
				end if
			end if
		end if
		if len(strMAge) > 0 then
			if not isNumeric(strMAge) then
				Err_Msg = Err_Msg & "<li>You must enter a numerical value for your age.</li>"
			elseif strMinAge > 0 and CInt(strMAge) < strMinAge then
				Err_Msg = Err_Msg & "<li>You must be at least " & strMinAge & " years old.</li>"
			end if
		end if
		if strAgeDOB = "1" and strReqAgeDOB = "1" then
			if len(strMDOB) <> 8 then
				Err_Msg = Err_Msg & "<li>You Must Provide Your Date Of Birth</li>"
			end if
 		end if
		if Err_Msg = "" then
			if Trim(Request.Form("Homepage")) <> "" and lcase(trim(Request.Form("Homepage"))) <> "http://" and Trim(lcase(Request.Form("Homepage"))) <> "https://" then
				regHomepage = ChkString(Request.Form("Homepage"),"SQLString")
			else
				regHomepage = " "
			end if
			if Trim(Request.Form("LINK1")) <> "" and lcase(trim(Request.Form("LINK1"))) <> "http://" and Trim(lcase(Request.Form("LINK1"))) <> "https://" then
				regLink1 = ChkString(Request.Form("LINK1"),"SQLString")
			else
				regLink1 = " "
			end if
			if Trim(Request.Form("LINK2")) <> "" and lcase(trim(Request.Form("LINK2"))) <> "http://" and Trim(lcase(Request.Form("LINK2"))) <> "https://" then
				regLink2 = ChkString(Request.Form("LINK2"),"SQLString")
			else
				regLink2 = " "
			end if
			if Trim(Request.Form("Photo_URL")) <> "" and lcase(trim(Request.Form("Photo_URL"))) <> "http://" and Trim(lcase(Request.Form("Photo_URL"))) <> "https://" then
				regPhoto_URL = ChkString(Request.Form("Photo_URL"),"SQLString")
			else
				regPhoto_URL = " "
			end if

			'## Forum_SQL
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
			if trim(Request.Form("Password")) <> "" then
				strPassword = sha256("" & Request.Form("Password"))
				strSql = strSql & " SET M_PASSWORD = '" & ChkString(strPassword,"SQLString") & "', "
			else
				strSql = strSql & " SET"
			end if
			strSql = strSql & "     M_COUNTRY  = '" & ChkString(Request.Form("Country"),"SQLString")  & "', "
			if strAIM = "1" then
				strSql = strSql & "     M_AIM = '" & ChkString(Request.Form("AIM"),"SQLString") & "', "
			end if
			if strICQ = "1" then
				strSql = strSql & "     M_ICQ = '" & ChkString(Request.Form("ICQ"),"SQLString") & "', "
			end if
			if strMSN = "1" then
				strSql = strSql & "     M_MSN = '" & ChkString(Request.Form("MSN"),"SQLString") & "', "
			end if
			if strYAHOO = "1" then
				strSql = strSql & "     M_YAHOO = '" & ChkString(Request.Form("YAHOO"),"SQLString") & "', "
			end if
			if strHOMEPAGE = "1" then
				strSql = strSql & "     M_HOMEPAGE = '" & ChkString(Trim(regHomepage),"SQLString") & "', "
			end if
			if strSignatures = "1" then
				strSql = strSql & "     M_SIG = '" & ChkString(Request.Form("Sig"),"message") & "', "
			end if
			if strSignatures = "1" and strDSignatures = "1" then
				strSql = strSql & "     M_VIEW_SIG = " & cLng(Request.Form("ViewSig")) & ", "
			end if
			if strSignatures = "1" then
				strSql = strSql & "     M_SIG_DEFAULT = " & cLng(Request.Form("fSigDefault")) & ", "
			end if
			if strEmailVal = "1" then
				strSql = strSql & "     M_NEWEMAIL = '" & ChkString(Request.Form("Email"),"SQLString") & "' "
			else
				strSql = strSql & "     M_EMAIL = '" & ChkString(Request.Form("Email"),"SQLString") & "' "
			end if
			strSql = strSql & ", 	M_KEY = '" & chkString(verKey,"SQLString") & "'"
			strSql = strSql & ",     M_RECEIVE_EMAIL = " & cLng(Request.Form("ReceiveEMail")) & " "
			if strfullName = "1" then
				strSql = strSql & ",	M_FIRSTNAME = '" & ChkString(Request.Form("FirstName"), "SQLString") & "'"
				strSql = strSql & ",	M_LASTNAME  = '" & ChkString(Request.Form("LastName"),"SQLString") & "'"
			end if
			if strCity = "1" then
				strsql = strsql & ",	M_CITY = '" & ChkString(Request.Form("City"),"SQLString") & "'"
			end if
			if strState = "1" then
				strsql = strsql & ",	M_STATE = '" & ChkString(Request.Form("State"),"SQLString") & "'"
			end if
'			strsql = strsql & ",	M_HIDE_EMAIL = '" & ChkString(Request.Form("HideMail"),"SQLString") & "'"
			if strPicture = "1" then
				strsql = strsql & ",	M_PHOTO_URL = '" & ChkString(Trim(regPhoto_URL),"SQLString") & "'"
			end if
			if strFavLinks = "1" then
				strsql = strsql & ",	M_LINK1 = '" & ChkString(Trim(regLink1),"SQLString") & "'"
				strSql = strSql & ",	M_LINK2 = '" & ChkString(Trim(regLink2),"SQLString") & "'"
			end if
			if strAge = "1" then
				strSql = strsql & ",	M_AGE = '" & strMAge & "'"
			end if
			if strAgeDOB = "1" then
				strSql = strsql & ",	M_DOB = '" & strMDOB & "'"
			end if
			if strMarStatus = "1" then
				strSql = strSql & ",	M_MARSTATUS = '" & ChkString(Request.Form("MarStatus"),"SQLString") & "'"
			end if
			if strSex = "1" then
				strSql = strsql & ",	M_SEX = '" & ChkString(Request.Form("Sex"),"SQLString") & "'"
			end if
			if strOccupation = "1" then
				strSql = strSql & ",	M_OCCUPATION = '" & ChkString(Request.Form("Occupation"),"SQLString") & "'"
			end if
			if strHobbies = "1" then
				strSql = strSql & ",	M_HOBBIES = '" & ChkString(Request.Form("Hobbies"),"message") & "'"
			end if
			if strQuote = "1" then
				strSql = strSql & ",	M_QUOTE = '" & ChkString(Request.Form("Quote"),"message") & "'"
			end if
			if strLNews = "1" then
				strsql = strsql & ",	M_LNEWS = '" & ChkString(Request.Form("LNews"),"message") & "'"
			end if
			if strBio = "1" then
				strSql = strSql & ",	M_BIO = '" & ChkString(Request.Form("Bio"),"message") & "'"
			end if
			strSql = strSql & " WHERE M_NAME = '" & ChkString(Request.Form("Name"), "SQLString") & "' "
			if strAuthType = "db" then 
				strSql = strSql & " AND   M_PASSWORD = '" & ChkString(Request.Form("Password-d"), "SQLString") & "'"
			end if

			my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			
			regHomepage = ""

			if trim(Request.Form("Password")) <> "" and strDBNTUserName <> "" then
				if strSetCookieToForum = 1 then
					Response.Cookies(strUniqueID & "User").Path = strCookieURL
				else
					Response.Cookies(strUniqueID & "User").Path = "/"
				end if
				Response.Cookies(strUniqueID & "User")("Pword") = strPassword
				Response.Cookies(strUniqueID & "User").Expires = dateAdd("d", intCookieDuration, strForumTimeAdjust)
			end if

			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Profile Updated.</font></p>" & vbNewLine
			if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) and lcase(strEmail) = "1" and strEmailVal = "1" then
				if (strUseExtendedProfile) then
					Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your e-mail address has changed. To complete your e-mail address change,<br />please follow the instructions in the e-mail that has been sent to your new e-mail address.</font></p>" & vbNewLine & _
									"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href="""
					if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(chkString(Request.Form("refer"),"refer"))
					Response.Write	""">Back To Forum</a></font></p>" & vbNewLine
				else
					Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your e-mail address has changed. To complete your e-mail address change, please follow the instructions in the e-mail that has been sent to your new e-mail address.<br /><br /></font></p>" & vbNewLine
				end if
			else
				if (strUseExtendedProfile) then
					Response.Write	"<meta http-equiv=""Refresh"" content=""2; URL=" & strRefer & """>" & vbNewLine & _
									"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
				end if
			end if
		else
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
							"<table align=""center"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
			end if 
		end if

	case "ModifyIt"
		if strUseExtendedProfile then
			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
							"" & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
							"" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Modify Profile</font></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine
		end if
		strEncodedPassword = ChkString(Request.Form("Pass"),"SQLString")
		mLev = cLng(chkUser(strDBNTUserName, strEncodedPassword,-1)) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then '## is Admin

				Err_Msg = ""
			
				if trim(Request.Form("Name")) = "" then 
					Err_Msg = Err_Msg & "<li>You must set a UserName</li>"
				end if
				if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) then
					Err_Msg = Err_Msg & "<li> &gt; and &lt; are not allowed in the UserName, Please Choose Another</li>"
				end if

				'## Forum_SQL
				strSql = "SELECT M_NAME FROM " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " WHERE M_NAME = '" & Trim(ChkString(Request.Form("Name"), "SQLString")) &"' "
				strSql = strSql & " AND MEMBER_ID <> " & cLng(Request.Form("Member_ID")) &" "
		
				set rs = my_Conn.Execute(TopSQL(strSql,1))	

				if rs.BOF and rs.EOF then 
					'## Do Nothing - proceed
				else 
					Err_Msg = Err_Msg & "<li>UserName is already in use, <br />Please Choose Another</li>"
				end if
						
				set rs = nothing

				if strEmail = "1" and strEmailVal = "1" then
					'## Forum_SQL
					strSql = "SELECT M_NAME FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
					strSql = strSql & " WHERE M_NAME = '" & Trim(ChkString(Request.Form("Name"), "SQLString")) &"' "
					strSql = strSql & " AND MEMBER_ID <> " & cLng(Request.Form("Member_ID")) &" "

					set rs = my_Conn.Execute(TopSQL(strSql,1))

					if rs.BOF and rs.EOF then 
						'## Do Nothing
					else
						Err_Msg = Err_Msg & "<li>UserName is already in use, <br />Please Choose Another</li>"
					end if
					set rs = nothing
				end if
				if strAuthType = "db" then
					if trim(Request.Form("Password")) <> "" then 
						if Len(Request.Form("Password")) > 25 then 
							Err_Msg = Err_Msg & "<li>The Password can not be greater than 25 characters</li>" 
						end if
					end if
				end if
				If Request.Form("Email") = "" Then
					Err_Msg = Err_Msg & "<li>You Must give an e-mail address</li>"
				Else
					'Comment out down to the next comment to let it take me@example.com and/or .ex as well
					'strsql = "SELECT SPAM_SERVER FROM " & strTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '" & chkString(Request.Form("Email"),"sqlstring") & "'"
					'set rsSpam = my_Conn.Execute (strsql)
					
					'If Not rsSpam.EOF Then
					'	Err_Msg = Err_Msg & "<li>You cannot register with '" & chkString(Request.Form("Email"),"sqlstring") & "'.</li>"
					'End If 
					
					'Dim strMailTLD : strMailTLD = LCase(Mid(Request.Form("Email"),InStrRev(Request.Form("Email"),".")))
					
					'strsql = "SELECT SPAM_SERVER FROM " & strTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '" & strMailTLD & "'"
					'set rsSpam = my_Conn.Execute (strsql)
					
					'If Not rsSpam.EOF Then
					'	Err_Msg = Err_Msg & "<li>You cannot register with a '" & strMailTLD & "' email address.</li>"
					'End If 
					'Comment out up to the previous comment to let it take me@example.com and/or .ex as well
					
					strMailDomain = LCase(Mid(Request.Form("Email"),InStrRev(Request.Form("Email"),"@")))
					
					strsql = "SELECT SPAM_SERVER FROM " & strTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '" & strMailDomain & "'"
					set rsSpam = my_Conn.Execute (strsql)
					
					If Not rsSpam.EOF Then
						Err_Msg = Err_Msg & "<li>You cannot register with an '" & strMailDomain & "' email address.</li>"
					End If 
					
					rsSpam.close
					Set rsSpam = Nothing
				End If
				if EmailField(Request.Form("Email")) = 0 then 
					Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address</li>"
				end if
				if strMSN = "1" and trim(Request.Form("MSN")) <> "" then
					if EmailField(Request.Form("MSN")) = 0 then 
						Err_Msg = Err_Msg & "<li>You Must enter a valid MSN Messenger Username</li>"
					end if
				end if
				if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
					Err_Msg = Err_Msg & "<li>You Must prefix the URL with <b>http://</b> or <b>https://</b></li>"
				end if
				if strUniqueEmail = "1" then
					if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) then
						'## Forum_SQL
						strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
						strSql = strSql & " WHERE M_EMAIL = '" & Trim(chkString(Request.Form("Email"),"SQLString")) &"'"

						set rs = my_Conn.Execute(TopSQL(strSql,1))

						if rs.BOF and rs.EOF then
							'## Do Nothing - proceed
						Else
							Err_Msg = Err_Msg & "<li>E-mail Address already in use, Please Choose Another</li>"
						end if
						set rs = nothing

						if strEmail = "1" and strEmailVal = "1" then
							'## Forum_SQL
							strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
							strSql = strSql & " WHERE M_EMAIL = '" & Trim(chkString(Request.Form("Email"),"SQLString")) &"'"

							set rs = my_Conn.Execute(TopSQL(strSql,1))

							if rs.BOF and rs.EOF then 
								'## Do Nothing
							else
								Err_Msg = Err_Msg & "<li>E-mail Address already in use, Please Choose Another</li>"
							end if
							set rs = nothing

							'## Forum_SQL
							strSql = "SELECT M_NEWEMAIL FROM " & strMemberTablePrefix & "MEMBERS "
							strSql = strSql & " WHERE M_NEWEMAIL = '" & Trim(ChkString(Request.Form("Email"),"SQLString")) &"'"

							set rs = my_Conn.Execute(TopSQL(strSql,1))

							if rs.BOF and rs.EOF then 
								'## Do Nothing
							else
								Err_Msg = Err_Msg & "<li>E-mail Address already in use, Please Choose Another</li>"
							end if
							set rs = nothing
						end if
						if lcase(strEmail) = "1" and Err_Msg = "" and strEmailVal = "1" then
							verKey = GetKey("sendemail")
						end if
					end if
				else
					if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) and lcase(strEmail) = "1" and strEmailVal = "1" then
						verKey = GetKey("sendemail")
					end if
				end if
				if not IsValidURL(trim(Request.Form("Homepage"))) then
					Err_Msg = Err_Msg & "<li>Homepage URL: Invalid URL" & strURLError & "</li>"
				end if
				if not IsValidURL(trim(Request.Form("LINK1"))) then
					Err_Msg = Err_Msg & "<li>Cool Links URL: Invalid URL" & strURLError & "</li>"
				end if
				if not IsValidURL(trim(Request.Form("LINK2"))) then
					Err_Msg = Err_Msg & "<li>Cool Links URL: Invalid URL" & strURLError & "</li>"
				end if
				if not IsValidURL(trim(Request.Form("Photo_URL"))) then
					Err_Msg = Err_Msg & "<li>Photo URL: Invalid URL" & strURLError & "</li>"
				end if
				strMAge = ""
				if strAge = "1" then
					strMAge = ChkString(trim(Request.Form("Age")), "SQLString")
				end if
				if strAgeDOB = "1" then
					strMDOB = ChkString(Request.Form("year"), "SQLString") & ChkString(Request.Form("month"), "SQLString") & ChkString(Request.Form("day"), "SQLString")
					if len(strMDOB) <> 8 then
						strMDOB = ""
					else
						if not IsValidBirthDate() then
							Err_Msg = Err_Msg & "<li>Date of Birth: Invalid Date</li>"
						else
							strMAge = DisplayUsersAge(DOBToDate(strMDOB))
						end if
					end if
				end if
				if len(strMAge) > 0 then
					if not isNumeric(strMAge) then
						Err_Msg = Err_Msg & "<li>You must enter a numerical value for your age.</li>"
					elseif strMinAge > 0 and strMAge < strMinAge then
						Err_Msg = Err_Msg & "<li>You must be at least " & strMinAge & " years old.</li>"
					end if
				end if
				if Err_Msg = "" then '## it is ok to update the profile
					if Trim(Request.Form("Homepage")) <> "" and lcase(trim(Request.Form("Homepage"))) <> "http://" and Trim(lcase(Request.Form("Homepage"))) <> "https://" then
						regHomepage = chkString(Request.Form("Homepage"),"SQLString")
					else
						regHomepage = " "
					end if
					if Trim(Request.Form("LINK1")) <> "" and lcase(trim(Request.Form("LINK1"))) <> "http://" and Trim(lcase(Request.Form("LINK1"))) <> "https://" then
						regLink1 = chkString(Request.Form("LINK1"),"SQLString")
					else
						regLink1 = " "
					end if
					if Trim(Request.Form("LINK2")) <> "" and lcase(trim(Request.Form("LINK2"))) <> "http://" and Trim(lcase(Request.Form("LINK2"))) <> "https://" then
						regLink2 = chkString(Request.Form("LINK2"),"SQLString")
					else
						regLink2 = " "
					end if
					if Trim(Request.Form("PHOTO_URL")) <> "" and lcase(trim(Request.Form("PHOTO_URL"))) <> "http://" and Trim(lcase(Request.Form("PHOTO_URL"))) <> "https://" then
						regPhoto_URL = chkString(Request.Form("Photo_URL"),"SQLString")
					else
						regPhoto_URL = " "
					end if
			
					'## Forum_SQL
					strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " SET M_NAME = '" & chkString(Request.Form("Name"),"SQLString") & "'"
					if strAuthType = "nt" then
						strSql = strSql & ",    M_USERNAME = '" & chkString(Request.Form("Account"),"SQLString") & "'"
					else
						if strAuthType = "db" then
							if trim(Request.Form("Password")) <> "" then
								strPassword = sha256("" & Request.Form("Password"))
								strSql = strSql & ", M_PASSWORD = '" & ChkString(strPassword,"SQLString") & "' "
							end if
						end if
					end if
					if strEmailVal = "1" then
						strSql = strSql & ", M_NEWEMAIL = '" & chkString(Request.Form("Email"),"SQLString") & "'"
					else
						strSql = strSql & ", M_EMAIL = '" & chkString(Request.Form("Email"),"SQLString") & "'"
					end if
					strSql = strSql & ", M_KEY = '" & chkString(verKey,"SQLString") & "'"
					strSql = strSql & ", M_RECEIVE_EMAIL = " & cLng(Request.Form("ReceiveEMail")) & " "
					strSql = strSql & ", M_TITLE = '" & chkString(Request.Form("Title"),"SQLString") & "'"
					strSql = strSql & ", M_POSTS = " & cLng(Request.Form("Posts")) & " "
					strSql = strSql & ", M_COUNTRY = '" & chkString(Request.Form("Country"),"SQLString") & "'"
					if strAIM = "1" then
						strSql = strSql & ", M_AIM = '" & chkString(Request.Form("AIM"),"SQLString") & "'"
					end if
					if strICQ = "1" then
						strSql = strSql & ", M_ICQ = '" & chkString(Request.Form("ICQ"),"SQLString") & "'"
					end if
					if strMSN = "1" then
						strSql = strSql & ", M_MSN = '" & chkString(Request.Form("MSN"),"SQLString") & "'"
					end if
					if strYAHOO = "1" then
						strSql = strSql & ", M_YAHOO = '" & chkString(Request.Form("YAHOO"),"SQLString") & "'"
					end if
					if strHOMEPAGE = "1" then
						strSql = strSql & ", M_HOMEPAGE = '" & chkString(Trim(regHomepage),"SQLString") & "'"
					end if
					if strSignatures = "1" then
						strSql = strSql & ", M_SIG = '" & chkString(Request.Form("Sig"),"message") & "'"
					end if
					'if strSignatures = "1" and strDSignatures = "1" then
					'	strSql = strSql & ", M_VIEW_SIG = " & cLng("0" & Request.Form("ViewSig"))
					'end if
					'if strSignatures = "1" then
					'	strSql = strSql & ", M_SIG_DEFAULT = " & cLng("0" & Request.Form("fSigDefault"))
					'end if
					strSql = strSql & ", M_LEVEL = " & cLng("0" & Request.Form("Level"))
					if strfullName = "1" then
						strSql = strSql & ", M_FIRSTNAME = '" & chkString(Request.Form("FirstName"),"SQLString") & "'" 
						strSql = strSql & ", M_LASTNAME  = '" & chkString(Request.Form("LastName"),"SQLString") & "'"  
					end if
					if strCity = "1" then
						strsql = strsql & ", M_CITY = '" & chkString(Request.Form("City"),"SQLString") & "'"  
					end if
					if strState = "1" then
						strsql = strsql & ", M_STATE = '" & chkString(Request.Form("State"),"SQLString") & "'" 
					end if
'					strsql = strsql & ",	M_HIDE_EMAIL = '" & chkString(Request.Form("HideMail"),"SQLString") & "'"  
					if strPicture = "1" then
						strsql = strsql & ", M_PHOTO_URL = '" & chkString(Trim(regPhoto_URL),"SQLString") & "'"  
					end if
					if strFavLinks = "1" then
						strsql = strsql & ", M_LINK1 = '" & chkString(Trim(regLink1),"SQLString") & "'" 
						strSql = strSql & ", M_LINK2 = '" & chkString(Trim(regLink2),"SQLString") & "'" 
					end if
					if strAge = "1" then
						strSql = strsql & ", M_AGE = '" & strMAge & "'" 
					end if
					if strAgeDOB = "1" then
						strSql = strsql & ", M_DOB = '" & strMDOB & "'"
					end if
					if strMarStatus = "1" then
						strSql = strSql & ", M_MARSTATUS = '" & chkString(Request.Form("MarStatus"),"SQLString") & "'" 
					end if
					if strSex = "1" then
						strSql = strsql & ", M_SEX = '" & chkString(Request.Form("Sex"),"SQLString") & "'" 
					end if
					if strOccupation = "1" then
						strSql = strSql & ", M_OCCUPATION = '" & chkString(Request.Form("Occupation"),"SQLString") & "'" 
					end if
					if strHobbies = "1" then
						strSql = strSql & ", M_HOBBIES = '" & chkString(Request.Form("Hobbies"),"message") & "'" 
					end if
					if strQuote = "1" then
						strSql = strSql & ", M_QUOTE = '" & chkString(Request.Form("Quote"),"message") & "'" 
					end if
					if strLNews = "1" then
						strsql = strsql & ", M_LNEWS = '" & chkString(Request.Form("LNews"),"message") & "'" 
					end if
					if strBio = "1" then
						strSql = strSql & ", M_BIO = '" & chkString(Request.Form("Bio"),"message") & "'" 
					end if
					If not IsNumeric(Request.Form("allowemail")) Then
						strSQL = strSql & ", M_ALLOWEMAIL = " & 0
					Else
						If Request.Form("allowemail") <> "1" Then
							strSQL = strSql & ", M_ALLOWEMAIL = " & 0
						Else
							strSQL = strSql & ", M_ALLOWEMAIL = " & cLng(Request.Form("allowemail"))
						End If
					End If
					strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))

					my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
					
					if ChkString(Request.Form("Level"),"") = "1" then 
						'## Forum_SQL - Remove the member from the moderator table
						strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
						strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					end if
					
					'## delete member from existing groups
					strSql = "DELETE FROM " & strTablePrefix & "USERGROUP_USERS " &_
						"WHERE MEMBER_TYPE = 1 AND MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))
					my_Conn.Execute(strSql)
					'## add member to new groups
					strGroups = Request.Form("UserGroups")
					if strGroups <> "" or not IsEmpty(strGroups) or not IsNull(strGroups) then
						arAddGroups = Split(strGroups,",")
						for iAGcount = LBound(arAddGroups) to UBound(arAddGroups)
							strSql = "INSERT INTO " & strTablePrefix & "USERGROUP_USERS " &_
								"(USERGROUP_ID, MEMBER_TYPE, MEMBER_ID) VALUES " &_
								"(" & cLng(Trim(arAddGroups(iAGcount))) & ", 1, " & cLng(Request.Form("MEMBER_ID")) & ")"
							my_Conn.Execute(strSql)
						next
					end if
					
					Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Profile Updated.</font></p>" & vbNewLine
					if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) and lcase(strEmail) = "1" and strEmailVal = "1" then
						if (strUseExtendedProfile) then
							Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>The e-mail address has been changed. A confirmation has been sent to the new e-mail address.</font></p>" & vbNewLine & _
											"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href="""
							if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(chkString(Request.Form("refer"),"refer"))
							Response.Write	""">Back To Forum</a></font></p>" & vbNewLine
						else
							Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>The e-mail address has been changed. A confirmation has been sent to the new e-mail address.<br /><br /></font></p>" & vbNewLine
						end if
					else
						if (strUseExtendedProfile) then
							Response.Write	"<meta http-equiv=""Refresh"" content=""2; URL=" & strRefer & """>" & vbNewLine & _
											"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
						end if
					end if
				else
					Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With The Details</font></p>" & vbNewLine & _
									"<table align=""center"">" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"</table>" & vbNewLine & _
									"<p align=""center""><font size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
					if strUseExtendedProfile then
						Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
					end if 
				end if
			else 'Member but no Admin
				Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Modify a Member</b></font><br />" & vbNewLine & _
								"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
				if strUseExtendedProfile then
					Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
				end if 
			end if
		else  'Not logged on or no member
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Modify a Member</b></font><br />" & vbNewLine & _
							"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strRefer & """>Back To Forum</a></font></p>" & vbNewLine
			end if 
		end if 
	case else	
		Response.Redirect("default.asp")
end select

set rs = nothing
if not(strUseExtendedProfile) then
	WriteFooterShort
	Response.End
else 
	WriteFooter
	Response.End
end if

function IsValidBirthDate()
	IsValidBirthDate = true
	strMDOByear = cInt(left(strMDOB, 4))
	strMDOBmonth = cInt(mid(strMDOB, 5, 2))
	strMDOBday = cInt(right(strMDOB, 2))
	arrDays = array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
	intDays = arrDays(strMDOBMonth - 1)
	if strMDOBmonth = 2 and strMDOByear mod 4 = 0 and not (strMDOByear mod 100 = 0 and not strMDOBYear mod 400 = 0) then
		intDays = intDays + 1
	end if
	if strMDOBday > intDays or strMDOB > left(DateToStr(strForumTimeAdjust), 8) then
		IsValidBirthDate = false
	end if
end function

Function IsValidURL(sValidate)
	Dim sInvalidChars
	Dim bTemp
	Dim i

	if trim(sValidate) = "" then IsValidURL = true : exit function
	sInvalidChars = """;+()*'<>"
	for i = 1 To Len(sInvalidChars)
		if InStr(sValidate, Mid(sInvalidChars, i, 1)) > 0 then bTemp = True
		if bTemp then strURLError = "<br />&bull;&nbsp;cannot contain any of the following characters:  "" ; + ( ) * ' < > "
		if bTemp then Exit For
	next
	if not bTemp then
		for i = 1 to Len(sValidate)
			if Asc(Mid(sValidate, i, 1)) = 160 then bTemp = True
			if bTemp then strURLError = "<br />&bull;&nbsp;cannot contain any spaces "
			if bTemp then Exit For
		next
	end if

	' extra checks
	' check to make sure URL begins with http:// or https://
	if not bTemp then
		bTemp = (lcase(left(sValidate, 7)) <> "http://") and (lcase(left(sValidate, 8)) <> "https://")
		if bTemp then strURLError = "<br />&bull;&nbsp;must begin with either http:// or https:// "
	end if
	' check to make sure URL is 255 characters or less
	if not bTemp then
		bTemp = len(sValidate) > 255
		if bTemp then strURLError = "<br />&bull;&nbsp;cannot be more than 255 characters "
	end if
	' no two consecutive dots
	if not bTemp then
		bTemp = InStr(sValidate, "..") > 0
		if bTemp then strURLError = "<br />&bull;&nbsp;cannot contain consecutive periods "
	end if
	'no spaces
	if not bTemp then
		bTemp = InStr(sValidate, " ") > 0
		if bTemp then strURLError = "<br />&bull;&nbsp;cannot contain any spaces "
	end if
	if not bTemp then
		bTemp = (len(sValidate) <> len(Trim(sValidate)))
		if bTemp then strURLError = "<br />&bull;&nbsp;cannot contain any spaces "
	end if 'Addition for leading and trailing spaces

	' if any of the above are true, invalid string
	IsValidURL = Not bTemp
End Function
%>
