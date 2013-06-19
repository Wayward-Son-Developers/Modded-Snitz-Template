<%
'##################################################################################################
'## Snitz Forums 2000 v3.4.07
'##################################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##		   Huw Reddick and Richard Kinser
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
'##################################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_profile.asp" -->
<!--#include FILE="inc_func_posting.asp"-->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%
Dim strURLError

if Instr(1,Request.Form("refer"),"search.asp",1) > 0 then
	strRefer = "search.asp"
elseif Instr(1,Request.Form("refer"),"register.asp",1) > 0 then	
	strRefer = "default.asp"
else	
	strRefer = Request.Form("refer")
end if
if strRefer = "" then strRefer = "default.asp"

if Request.QueryString("id") <> "" and IsNumeric(Request.QueryString("id")) = true then
	ppMember_ID = cLng(Request.QueryString("id"))
else
	ppMember_ID = 0
end if

if strAuthType = "nt" then
	if ChkAccountReg() <> "1" then 
		Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>" & vbNewLine & _
				"<b>Note:</b> This NT account has not been registered yet, thus the profile is not available.<br />" & vbNewLine
		if strProhibitNewMembers <> "1" then
			Response.Write	"If this is your account, <a href=""policy.asp"">click here</a> to register.</font></p>" & vbNewLine
		else
			Response.Write	"</font></p>" & vbNewLine
		end if
 		WriteFooter
		Response.End 
	end if
end if

select case Request.QueryString("mode") 

	case "display" '## Display Profile

		if strDBNTUserName = "" then
			Err_Msg = "You must be logged in to view a Member's Profile"

			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
				"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_accounts_pending.asp"">Members&nbsp;Pending</a><br />" & vbNewLine & _
				"		" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member's Profile</font></td>" & vbNewLine & _
				"	</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem!</font></p>" & vbNewLine & _
				"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>" & Err_Msg & "</font></p>" & vbNewLine & _
				"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Back to Forum</a></font></p>" & vbNewLine & _
				"	<br />" & vbNewLine
				WriteFooterShort
				Response.End
		end if

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS_PENDING.MEMBER_ID"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_NAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_USERNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_FIRSTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LASTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_MSN"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_COUNTRY"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_CITY"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_STATE"
'		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_HIDE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_RECEIVE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_DATE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_PHOTO_URL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_HOMEPAGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LINK1"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LINK2"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_AGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_DOB"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_MARSTATUS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_SEX"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_OCCUPATION"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_HOBBIES"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_QUOTE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LNEWS"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_BIO"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
		strSql = strSql & " WHERE MEMBER_ID=" & ppMember_ID

		set rs = my_Conn.Execute(strSql)
		
		if rs.BOF or rs.EOF then
			Err_Msg = "Invalid Member ID!"

			Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
				"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_accounts_pending.asp"">Members&nbsp;Pending</a><br />" & vbNewLine & _
				"		" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Member's Profile</font></td>" & vbNewLine & _
				"	</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem!</font></p>" & vbNewLine & _
				"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>" & Err_Msg & "</font></p>" & vbNewLine & _
				"	<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Back to Forum</a></font></p>" & vbNewLine & _
				"	<br />" & vbNewLine
				WriteFooter
				Response.End
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
			else
				strColspan = ""
				strIMURL1 = ""
				strIMURL2 = ""
			end if

			if strUseExtendedProfile then
				Response.Write	"<table width=""100%"" border=""0"">" & vbNewLine & _
					"	<tr>" & vbNewLine & _
					"		<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
					"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
					"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
					"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_accounts_pending.asp"">Members&nbsp;Pending</a><br />" & vbNewLine & _
					"		" & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;" & chkString(rs("M_NAME"),"display") & "'s Profile</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine & _
					"</table>" & vbNewLine
			end if
			Response.Write	"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td bgColor=""" & strPageBGColor & """ align=""center""" & strColspan & ">" & vbNewLine & _
				"		<font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Pending Profile<br /></font></td>" & vbNewLine & _
				"	</tr>" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td bgColor=""" & strPageBGColor & """ align=""center""" & strColspan & ">" & vbNewLine & _
				"<table border=""0"" width=""90%"" cellspacing=""0"" cellpadding=""4"" align=""center"">" & vbNewLine & _
				"	<tr>" & vbNewLine
			if mLev = 4 then
				Response.Write	"		<td valign=""top"" align=""left"" bgcolor=""" & strHeadCellColor & """>&nbsp;<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>" & ChkString(rs("M_NAME"),"display") & "</b></font></td>" & vbNewLine
			else
				Response.Write	"		<td valign=""top"" align=""left"" bgcolor=""" & strHeadCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>&nbsp;" & ChkString(rs("M_NAME"),"display") & "</b></font></td>" & vbNewLine
			end if
			Response.Write	"		<td valign=""top"" align=""right"" bgcolor=""" & strHeadCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Pending Since:&nbsp;" & ChkDate(rs("M_DATE"),"",True) & "&nbsp;(" & DateDiff("d",  StrToDate(rs("M_DATE")),  strForumTimeAdjust) & " days)</font></td>" & vbNewLine & _
				"	</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"		</td>" & vbNewLine & _
				"	</tr>" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td bgcolor=""" & strPageBGColor & """ align=""left"" valign=""top"">" & vbNewLine & _
				"<table border=""0"" width=""90%"" cellspacing=""1"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"	<tr>" & vbNewLine
			if strUseExtendedProfile then
				Response.Write	"		<td width=""35%"" bgColor=""" & strPageBGColor & """ valign=""top"">" & vbNewLine & _
					"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""3"">" & vbNewLine
				if trim(rs("M_PHOTO_URL")) = "" or lcase(rs("M_PHOTO_URL")) = "http://" then strPicture = 0
				if strPicture = "1" then
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>&nbsp;My Picture&nbsp;</font></b></td>" & vbNewLine & _
						"	</tr>" & vbNewLine & _
						"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2"">"
					if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then
						Response.Write	"<a href=""" & ChkString(rs("M_PHOTO_URL"), "displayimage") & """>" & getCurrentIcon(ChkString(rs("M_PHOTO_URL"), "displayimage") & "|150|150",ChkString(rs("M_NAME"),"display"),"hspace=""2"" vspace=""2""") & "</a><br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Click image for full picture</font>"
					else
						Response.Write	getCurrentIcon(strIconPhotoNone,"No Photo Available","hspace=""2"" vspace=""2""")
					end if
					Response.Write	"		</td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if ' strPicture
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>&nbsp;My Contact Info&nbsp;</font></b></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
				strContacts = 0
				if mLev > 2 or rs("M_RECEIVE_EMAIL") = "1" then
					strContacts = strContacts + 1
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" width=""10%"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>E-mail Address:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_EMAIL")) <> "" then
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & trim(rs("M_EMAIL")) & "</font></td>" & vbNewLine
					else
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No address specified...</font></td>" & vbNewLine
					end if
					Response.Write	"	</tr>" & vbNewLine
				end if
				if strAIM = "1" and Trim(rs("M_AIM")) <> "" then 
					strContacts = strContacts + 1
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>AIM:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconAIM,"","align=""absmiddle""") & "&nbsp;<a href=""" & strIMURL1 & "pop_messengers.asp?mode=AIM&ID=" & rs("MEMBER_ID") & strIMURL2 & """>" & ChkString(rs("M_AIM"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if 
				if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then 
					strContacts = strContacts + 1
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>ICQ:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & ChkString(rs("M_ICQ"), "urlpath") & "&img=5|18|18","","align=""absmiddle""") & "&nbsp;<a href=""" & strIMURL1 & "pop_messengers.asp?mode=ICQ&ID=" & rs("MEMBER_ID") & strIMURL2 & """>" & ChkString(rs("M_ICQ"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strMSN = "1" and Trim(rs("M_MSN")) <> "" then 
					strContacts = strContacts + 1
					parts = split(rs("M_MSN"),"@")
					strtag1 = parts(0)
					partss = split(parts(1),".")
					strtag2 = partss(0)
					strtag3 = partss(1)

					Response.Write	"		<script language=""javascript"" type=""text/javascript"">" & vbNewLine & _
						"			function MSNjs() {" & vbNewLine & _
						"			var tag1 = '" & strtag1 & "';" & vbNewLine & _
						"			var tag2 = '" & strtag2 & "';" & vbNewLine & _
						"			var tag3 = '" & strtag3 & "';" & vbNewLine & _
						"			document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & vbNewLine & _
						"		</script>" & vbNewLine
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>MSN:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconMSNM,"","align=""absmiddle""") & "&nbsp;<script language=""javascript"" type=""text/javascript"">MSNjs()</script>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then 
					strContacts = strContacts + 1
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>YAHOO IM:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(rs("M_YAHOO"), "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon("http://opi.yahoo.com/online?u=" & ChkString(rs("M_YAHOO"), "urlpath") & "&m=g&t=2|125|25","","") & "</a>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strContacts = 0 then
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No info specified...</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if


				if (strHomepage + strFavLinks) > 0 then  

					Response.Write	"	<tr>" & vbNewLine & _
						"		<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2"">" & vbNewLine & _
						"		<b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Links&nbsp;</font></b></td>" & vbNewLine
					if strHomepage = "1" then
						Response.Write	"	<tr>" & vbNewLine & _
							"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Homepage:&nbsp;</font></b></td>" & vbNewLine
						if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then
							Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & rs("M_HOMEPAGE") & """ target=""_blank"">" & rs("M_HOMEPAGE") & "</a>&nbsp;</font></td>" & vbNewLine
						else
							Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No homepage specified...</font></td>" & vbNewLine
						end if
						Response.Write	"	</tr>" & vbNewLine
					end if
					if strFavLinks = "1" then 
						Response.Write	"	<tr>" & vbNewLine & _
							"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Cool Links:&nbsp;</font></b></td>" & vbNewLine
						if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then
							Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & rs("M_LINK1") & """ target=""_blank"">" & rs("M_LINK1") & "</a>&nbsp;</font></td>" & vbNewLine
							if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then
								Response.Write	"	</tr>" & vbNewLine & _
									"	<tr>" & vbNewLine & _
									"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></b></td>" & vbNewLine & _
									"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & rs("M_LINK2") & """ target=""_blank"">" & rs("M_LINK2") & "</a>&nbsp;</font></td>" & vbNewLine
							end if
						else
							Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No link specified...</font></td>" & vbNewLine
						end if 
						Response.Write	"	</tr>" & vbNewLine
					end if
				end if ' strRecentTopics
				Response.Write	"</table>" & vbNewLine & _
					"		</td>" & vbNewLine & _
					"		<td valign=""top"" width=""3%"" bgColor=""" & strPageBGColor & """>&nbsp;</td>" & vbNewLine
			end if ' UseExtendedMemberProfile
			Response.Write	"		<td bgColor=""" & strPageBGColor & """ valign=""top"">" & vbNewLine & _
				"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""3"" valign=""top"">" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td valign=""top"" align=""center"" colspan=""2"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Basics</font></b></td>" & vbNewLine & _
				"	</tr>" & vbNewLine & _
				"	<tr>" & vbNewLine & _
				"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%"" valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:&nbsp;</font></b></td>" & vbNewLine & _
				"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_NAME"),"display") & "&nbsp;</font></td>" & vbNewLine & _
				"	</tr>" & vbNewLine
			if strAuthType = "nt" then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Account:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_USERNAME"),"display") & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if strFullName = "1" and (Trim(rs("M_FIRSTNAME")) <> "" or Trim(rs("M_LASTNAME")) <> "" ) then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Real Name:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_FIRSTNAME"), "display") & "&nbsp;" & ChkString(rs("M_LASTNAME"), "display") & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if (strCity = "1" and Trim(rs("M_CITY")) <> "") or (strCountry = "1" and Trim(rs("M_COUNTRY")) <> "") or (strState = "1" and Trim(rs("M_STATE")) <> "") then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Location:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
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
					"	</tr>" & vbNewLine
			end if
			if (strAge = "1" and Trim(rs("M_AGE")) <> "") then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Age:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_AGE"), "display") & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			strDOB = rs("M_DOB")
			if (strAgeDOB = "1" and Trim(strDOB) <> "") then
			strDOB = DOBToDate(strDOB)
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Age:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & DisplayUsersAge(strDOB) & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if (strMarStatus = "1" and Trim(rs("M_MARSTATUS")) <> "") then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Marital Status:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_MARSTATUS"), "display") & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if (strSex = "1" and Trim(rs("M_SEX")) <> "") then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Gender:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_SEX"), "display") & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if (strOccupation = "1" and Trim(rs("M_OCCUPATION")) <> "") then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Occupation:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(rs("M_OCCUPATION"), "display") & "</font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if intTotalMemberPosts > 0 then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Total Posts:&nbsp;</font></b></td>" & vbNewLine & _
					"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(intTotalMemberPosts, "display") & "<br /><font size=""" & strFooterFontSize & """>[" & strMemberPostsperDay & strPosts & " per day]<br /><a href=""search.asp?mode=DoIt&MEMBER_ID=" & rs("MEMBER_ID") & """>Find all non-archived posts by " & chkString(rs("M_NAME"),"display") & "</a></font></font></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
			end if
			if not(strUseExtendedProfile) then
				if rs("M_RECEIVE_EMAIL") = "1" then
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" width=""10%"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>E-mail Address:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_EMAIL")) <> "" then
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""pop_mail.asp?id=" & rs("MEMBER_ID") & """>Click to send an E-Mail</a>&nbsp;</font></td>" & vbNewLine
					else
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No address specified...</font></td>" & vbNewLine
					end if
					Response.Write	"	</tr>" & vbNewLine
				end if
				if strAIM = "1" and Trim(rs("M_AIM")) <> "" then
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>AIM:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconAIM,"","align=""absmiddle""") & "&nbsp;<a href=""pop_messengers.asp?mode=AIM&ID=" & rs("MEMBER_ID") & """>" & ChkString(rs("M_AIM"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if 
				if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then 
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>ICQ:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & ChkString(rs("M_ICQ"), "urlpath") & "&img=5|18|18","","align=""absmiddle""") & "&nbsp;<a href=""pop_messengers.asp?mode=ICQ&ID=" & rs("MEMBER_ID") & """>" & ChkString(rs("M_ICQ"), "display") & "</a>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strMSN = "1" and Trim(rs("M_MSN")) <> "" then
					parts = split(rs("M_MSN"),"@")
					strtag1 = parts(0)
					partss = split(parts(1),".")
					strtag2 = partss(0)
					strtag3 = partss(1)

					Response.Write	"<script language=""javascript"" type=""text/javascript"">" & vbNewLine & _
						"	function MSNjs() {" & vbNewLine & _
						"		var tag1 = '" & strtag1 & "';" & vbNewLine & _
						"		var tag2 = '" & strtag2 & "';" & vbNewLine & _
						"		var tag3 = '" & strtag3 & "';" & vbNewLine & _
						"		document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & vbNewLine & _
						"</script>" & vbNewLine

					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>MSN:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconMSNM,"","align=""absmiddle""") & "&nbsp;<script language=""javascript"" type=""text/javascript"">MSNjs()</script>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if 
				if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then 
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>YAHOO IM:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(rs("M_YAHOO"), "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon("http://opi.yahoo.com/online?u=" & ChkString(rs("M_YAHOO"), "urlpath") & "&m=g&t=2|125|25","","") & "</a>&nbsp;</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
			end if
			if IsNull(strMyBio) or trim(strMyBio) = "" then strBio = 0
			if IsNull(strMyHobbies) or trim(strMyHobbies) = "" then strHobbies = 0
			if IsNull(strMyLNews) or trim(strMyLNews) = "" then strLNews = 0
			if IsNull(strMyQuote) or trim(strMyQuote) = "" then strQuote = 0
			if (strBio + strHobbies + strLNews + strQuote) > 0 then
				Response.Write	"	<tr>" & vbNewLine & _
					"		<td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>More About Me</font></b></td>" & vbNewLine & _
					"	</tr>" & vbNewLine
				if strBio = "1" then
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Bio:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyBio) or trim(strMyBio) = "" then Response.Write("-") else Response.Write(formatStr(strMyBio))
					Response.Write	"</font></td>" & vbNewLine & _
						"		</tr>" & vbNewLine
				end if
				if strHobbies = "1" then  
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Hobbies:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyHobbies) or trim(strMyHobbies) = "" then Response.Write("-") else Response.Write(formatStr(strMyHobbies))
					Response.Write	"</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strLNews = "1" then  
					Response.Write	"		<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Latest News:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyLNews) or trim(strMyLNews) = "" then Response.Write("-") else Response.Write(formatStr(strMyLNews))
					Response.Write	"</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strQuote = "1" then  
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgcolor=""" & strPopUpTableColor & """ valign=""top"" align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Favorite Quote:&nbsp;</font></b></td>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if IsNull(strMyQuote) or Trim(strMyQuote) = "" then Response.Write("-") else Response.Write(formatStr(strMyQuote))
					Response.Write	"</font></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
			end if
			if (strHomepage + strFavLinks) > 0 and not(strRecentTopics = "0" and strUseExtendedProfile) then  
				if strUseExtendedProfile then	
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgcolor=""" & strCategoryCellColor & """ align=""center"" colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Links&nbsp;</font></b></td>" & vbNewLine & _
						"	</tr>" & vbNewLine
				end if
				if strHomepage = "1" then 
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Homepage:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & ChkString(rs("M_HOMEPAGE"), "display") & """ target=""_blank"">" & ChkString(rs("M_HOMEPAGE"), "display") & "</a>&nbsp;</font></td>" & vbNewLine
					else
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No homepage specified...</font></td>" & vbNewLine
					end if
					Response.Write	"	</tr>" & vbNewLine
				end if
				if strFavLinks = "1" then
					Response.Write	"	<tr>" & vbNewLine & _
						"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Cool Links:&nbsp;</font></b></td>" & vbNewLine
					if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then 
						Response.Write	"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & ChkString(rs("M_LINK1"), "display") & """ target=""_blank"">" & ChkString(rs("M_LINK1"), "display") & "</a>&nbsp;</font></td>" & vbNewLine
						if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then
							Response.Write	"	</tr>" & vbNewLine & _
								"	<tr>" & vbNewLine & _
								"		<td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap width=""10%""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></b></td>" & vbNewLine & _
								"		<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & ChkString(rs("M_LINK2"), "display") & """ target=""_blank"">" & ChkString(rs("M_LINK2"), "display") & "</a>&nbsp;</font></td>" & vbNewLine
						end if
					else
						Response.Write	"								<td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No link specified...</font></td>" & vbNewLine
					end if 
					Response.Write	"							</tr>" & vbNewLine
				end if
			end if
			Response.Write	"							</table>" & vbNewLine & _
					"								</td>" & vbNewLine & _
					"							</tr>" & vbNewLine & _
					"						</table>" & vbNewLine & _
					"					</td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table><br />" & vbNewLine & _
					"    </td>" & vbNewLine & _
					"  </tr>" & vbNewLine
				Response.Write	"  <tr>" & vbNewLine & _
						"    <td bgColor=""" & strPageBGColor & """ align=""center"" nowrap>" & vbNewLine
		end if
	case else	
		Response.Redirect("default.asp")
end select

set rs = nothing
	WriteFooter

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