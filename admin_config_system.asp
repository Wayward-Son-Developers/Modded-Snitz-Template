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
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Main&nbsp;Forum&nbsp;Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Request.Form("strTitleImage") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the Address of a Title Image</li>"
	end if
	if Request.Form("strHomeURL") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the URL of your HomePage (either relative or full)</li>"
	end if
	if Request.Form("strForumURL") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the Fully Qualified URL of your Forum</li>"
	end if
	if (left(lcase(Request.Form("strForumURL")), 7) <> "http://" and left(lcase(Request.Form("strForumURL")), 8) <> "https://") and Request.Form("strHomeURL") <> "" then
		Err_Msg = Err_Msg & "<li>You Must prefix the Forum URL with <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
	end if
	if (right(lcase(Request.Form("strForumURL")), 1) <> "/") then
		Err_Msg = Err_Msg & "<li>You Must end the Forum URL with <b>/</b></li>"
	end if
	if trim(Request.Form("strImageURL")) <> "" then
		if (right(lcase(Request.Form("strImageURL")), 1) <> "/") then
			Err_Msg = Err_Msg & "<li>You Must end the Images Location with <b>/</b></li>"
		end if
	end if
	if Request.Form("strAuthType") <> strAuthType and strAuthType = "db" then 
		if not(mLev = 4 and MemberID = intAdminMemberID) then
			Err_Msg = Err_Msg & "<li>Only the Admin user can change the Authentication type of the board</li>"
		else
			call NTauthenticate()
			if Session(strCookieURL & "userid") = "" then
				Err_Msg = Err_Msg & "<li>You have to enable non-Anonymous access for the forum on the server first</li>"
			else	
				strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"

				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords			
				call NTauthenticate()
				call NTUser()	
			end if
		end if
	end if
	if (Request.Form("strAuthType") <> strAuthType) and strAuthType = "nt" then 
		if not(mLev = 4 and MemberID = intAdminMemberID) then
			Err_Msg = Err_Msg & "<li>Only the Admin user can change the Authentication type of the board</li>"
		else
			Session(strCookieURL & "Approval") = "" 
		end if	
	end if
	if Err_Msg = "" then

		'## Forum_SQL
		for each key in Request.Form 
			if left(key,3) = "str" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next
		Application(strCookieURL & "ConfigLoaded") = ""
		Call OkMessage("Configuration Posted!","admin_home.asp","Back To Admin Home")
	else
		Call FailMessage(Err_Msg,True)
	end if
else
	Response.Write	"<form action=""admin_config_system.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Main Forum Configuration</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Forum's Title:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strForumTitle"" size=""30"" value=""" & chkExist(chkString(strForumTitle,"edit")) & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#forumtitle')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Forum's Copyright:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strCopyright"" size=""30"" value=""" & chkExistElse(chkString(strCopyright,"edit"),"2000-2002 Snitz Communications") & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#copyright')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Title Image Location:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strTitleImage"" size=""30"" value=""" & chkExistElse(strTitleImage,"logo_snitz_forums_2000.gif") & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#titleimage')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Home URL:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strHomeURL"" size=""30"" value=""" & chkExistElse(strHomeURL,"../") & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#homeurl')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Forum URL:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strForumURL"" size=""30"" value=""" & chkExistElse(strForumURL,"./") & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#forumurl')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Images Location:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<input type=""text"" name=""strImageURL"" size=""30"" value=""" & chkExist(strImageURL) & """>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#imagelocation')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Version info:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">"
	if strVersion <> "" then Response.Write("[<b>"& strVersion & "</b>]") else Response.Write("<b>[Couldn't read version info..]</b>")
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Authorization Type:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"DB: <input type=""radio"" class=""radio"" name=""strAuthType"" value=""db""" & chkRadio(strAuthType,"db",true) & ">" & vbNewLine & _
					"NT: <input type=""radio"" class=""radio"" name=""strAuthType"" value=""nt""" & chkRadio(strAuthType,"nt",true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#AuthType')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Set Cookie To:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"Forum: <input type=""radio"" class=""radio"" name=""strSetCookieToForum"" value=""1""" & chkRadio(strSetCookieToForum,1,true) & ">" & vbNewLine & _
					"WebSite: <input type=""radio"" class=""radio"" name=""strSetCookieToForum"" value=""0""" & chkRadio(strSetCookieToForum,0,true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#SetCookieToForum')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Use Graphics as Buttons:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strGfxButtons"" value=""1""" & chkRadio(strGfxButtons,1,true) & ">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strGfxButtons"" value=""0""" & chkRadio(strGfxButtons,0,true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#GfxButtons')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Use Graphic for ""Powered By"" link:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strShowImagePoweredBy"" value=""1""" & chkRadio(strShowImagePoweredBy,1,true) & ">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strShowImagePoweredBy"" value=""0""" & chkRadio(strShowImagePoweredBy,0,true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#PoweredBy')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Prohibit New Members:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strProhibitNewMembers"" value=""1""" & chkRadio(strProhibitNewMembers,1,true) & ">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strProhibitNewMembers"" value=""0""" & chkRadio(strProhibitNewMembers,0,true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#ProhibitNewMembers')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Require Registration:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strRequireReg"" value=""1""" & chkRadio(strRequireReg,1,true) & ">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strRequireReg"" value=""0""" & chkRadio(strRequireReg,0,true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#RequireReg')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">UserName Filter:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strUserNameFilter"" value=""1""" & chkRadio(strUserNameFilter,1,true) & ">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strUserNameFilter"" value=""0""" & chkRadio(strUserNameFilter,0,true) & ">" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#UserNameFilter')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"" id=""submit1"" name=""submit1"">Submit New Config</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if 
WriteFooter
Response.End
%>
