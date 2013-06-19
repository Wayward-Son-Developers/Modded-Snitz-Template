<%
'#################################################################################
'## Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or any later version.
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
'## Support can be obtained from support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## reinhold@bigfoot.com
'##
'## or
'##
'## Snitz Communications
'## C/O: Michael Anderson
'## PO Box 200
'## Harpswell, ME 04079
'#################################################################################
%>
<!--#INCLUDE FILE="inc_func_common.asp" -->
<%

if strShowTimer = "1" then
	'### start of timer code
	Dim StopWatch(19) 

	sub StartTimer(x)
		StopWatch(x) = timer
	end sub

	function StopTimer(x)
		EndTime = Timer

		'Watch for the midnight wraparound...
		if EndTime < StopWatch(x) then
			EndTime = EndTime + (86400)
		end if

		StopTimer = EndTime - StopWatch(x)
	end function

	StartTimer 1

	'### end of timer code
end if

strArchiveTablePrefix = strTablePrefix & "A_"
strScriptName = request.servervariables("script_name")

if Application(strCookieURL & "down") then 
	if not Instr(strScriptName,"admin_") > 0 then
		Response.redirect("down.asp")
	end if
end if

if strPageBGImageURL = "" then
	strTmpPageBGImageURL = ""
elseif Instr(strPageBGImageURL,"/") > 0 or Instr(strPageBGImageURL,"\") > 0 then
	strTmpPageBGImageURL = " background=""" & strPageBGImageURL & """"
else
	strTmpPageBGImageURL = " background=""" & strImageUrl & strPageBGImageURL & """"
end if

If strDBType = "" then 
	Response.Write	"<html>" & vbNewLine & _
			"<head>" & vbNewline & _
			"<title>" & strForumTitle & "</title>" & vbNewline


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

	Response.Write	"</head>" & vbNewLine & _
			"<body" & strTmpPageBGImageURL & " bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ link=""" & strLinkColor & """ aLink=""" & strActiveLinkColor & """ vLink=""" & strVisitedLinkColor & """>" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""40%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""navyblue"" align=""center""><p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
			"<b>There has been a problem...</b><br /><br />" & _
			"Your <b>strDBType</b> is not set, please edit your <b>config.asp</b><br />to reflect your database type." & _
			"</font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
			"<a href=""default.asp"" target=""_top"">Click here to retry.</a></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</body>" & vbNewLine & _
			"</html>" & vbNewLine
	Response.End
end if

set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

if (strAuthType = "nt") then
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

if strGroupCategories = "1" then
	if Request.QueryString("Group") = "" then
		if Request.Cookies(strCookieURL & "GROUP") = "" Then
			Group = 2
		else 
			Group = Request.Cookies(strCookieURL & "GROUP")
		end if
	else
		Group = cLng(Request.QueryString("Group"))
	end if
	'set default
	Session(strCookieURL & "GROUP_ICON") = "icon_group_categories.gif"
	Session(strCookieURL & "GROUP_IMAGE") = strTitleImage
	'Forum_SQL - Group exists ?
	strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_ICON, GROUP_IMAGE " 
	strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
	strSql = strSql & " WHERE GROUP_ID = " & Group
	set rs2 = my_Conn.Execute (strSql)
	if rs2.EOF or rs2.BOF then
		Group = 2
		strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_ICON, GROUP_IMAGE " 
		strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
		strSql = strSql & " WHERE GROUP_ID = " & Group
		set rs2 = my_Conn.Execute (strSql)
	end if	
	Session(strCookieURL & "GROUP_NAME") = rs2("GROUP_NAME")
	if instr(rs2("GROUP_ICON"), ".") then
		Session(strCookieURL & "GROUP_ICON") = rs2("GROUP_ICON")
	end if
	if instr(rs2("GROUP_IMAGE"), ".") then
		Session(strCookieURL & "GROUP_IMAGE") = rs2("GROUP_IMAGE")
	end if
	rs2.Close  
	set rs2 = nothing  
	Response.Cookies(strCookieURL & "GROUP") = Group
	Response.Cookies(strCookieURL & "GROUP").Expires =  dateAdd("d", intCookieDuration, strForumTimeAdjust)
	if Session(strCookieURL & "GROUP_IMAGE") <> "" then
		strTitleImage = Session(strCookieURL & "GROUP_IMAGE") 
	end if 
end if

strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
strDBNTFUserName = trim(chkString(Request.Form("Name"),"SQLString"))
if strDBNTFUserName = "" then strDBNTFUserName = trim(chkString(Request.Form("User"),"SQLString"))
if strAuthType = "nt" then
	strDBNTUserName = Session(strCookieURL & "userID")
	strDBNTFUserName = Session(strCookieURL & "userID")
end if

if strRequireReg = "1" and strDBNTUserName = "" then
	if not Instr(strScriptName,"policy.asp") > 0 and _
	not Instr(strScriptName,"register.asp") > 0 and _
	not Instr(strScriptName,"password.asp") > 0 and _
	not Instr(strScriptName,"faq.asp") > 0 and _
	not Instr(strScriptName,"login.asp") > 0 then
		scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
		if Request.QueryString <> "" then
			Response.Redirect("login.asp?target=" & lcase(scriptname(ubound(scriptname))) & "?" & Request.QueryString)
		else
			Response.Redirect("login.asp?target=" & lcase(scriptname(ubound(scriptname))))
		end if
	end if
end if

select case Request.Form("Method_Type")
	case "login"
		strEncodedPassword = sha256("" & Request.Form("Password"))
		select case chkUser(strDBNTFUserName, strEncodedPassword,-1)
			case 1, 2, 3, 4
				Call DoCookies(Request.Form("SavePassword"))
				strLoginStatus = 1
			case else
				strLoginStatus = 0
			end select
	case "logout"
		Call ClearCookies()
end select

if trim(strDBNTUserName) <> "" and trim(Request.Cookies(strUniqueID & "User")("Pword")) <> "" then
	chkCookie = 1
	mLev = cLng(chkUser(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"),-1))
	chkCookie = 0
else
	MemberID = -1
	mLev = 0
end if

if mLev = 4 and strEmailVal = "1" and strRestrictReg = "1" and strEmail = "1" then
	'## Forum_SQL - Get membercount from DB 
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE M_APPROVE = " & 0

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if not rs.EOF then
		User_Count = cLng(rs("U_COUNT"))
	else
		User_Count = 0
	end if

	rs.close
	set rs = nothing
end if

'#######    Begin UserGroup MOD     #######
if Session(strCookieURL & "UserGroups" & MemberID) = "" or _
	IsNull(Session(strCookieURL & "UserGroups" & MemberID)) then
	strGroupMembership = getGroupMembership(MemberID,1)
	Session(strCookieURL & "UserGroups" & MemberID) = strGroupMembership
	Session(strCookieURL & "UserGroups" & MemberID) = strGroupMembership
end if
'#######     End UserGroup MOD      #######

Response.Write	"<html>" & vbNewline & vbNewline & _
		"<head>" & vbNewline & _
		"<title>" & GetNewTitle(strScriptName) & "</title>" & vbNewline


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT



Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"<!-- hide from JavaScript-challenged browsers" & vbNewLine & _
		"function openWindow(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=400')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow2(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=450')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow3(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=450,scrollbars=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow4(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=525')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow5(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=450,height=525,scrollbars=yes,toolbars=yes,menubar=yes,resizable=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow6(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=500,height=450,scrollbars=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindowHelp(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=470,height=200,scrollbars=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"// done hiding -->" & vbNewLine & _
		"</script>" & vbNewLine & _
		"<style type=""text/css"">" & vbNewLine & _
		"<!--" & vbNewLine & _
		"a:link    {color:" & strLinkColor & ";text-decoration:" & strLinkTextDecoration & "}" & vbNewLine & _
		"a:visited {color:" & strVisitedLinkColor & ";text-decoration:" & strVisitedTextDecoration & "}" & vbNewLine & _
		"a:hover   {color:" & strHoverFontColor & ";text-decoration:" & strHoverTextDecoration & "}" & vbNewLine & _
		"a:active  {color:" & strActiveLinkColor & ";text-decoration:" & strActiveTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:link    {color:" & strForumLinkColor & ";text-decoration:" & strForumLinkTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:visited {color:" & strForumVisitedLinkColor & ";text-decoration:" & strForumVisitedTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:hover   {color:" & strForumHoverFontColor & ";text-decoration:" & strForumHoverTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:active  {color:" & strForumActiveLinkColor & ";text-decoration:" & strForumActiveTextDecoration & "}" & vbNewLine & _
		".spnSearchHighlight {background-color:" & strSearchHiLiteColor & "}" & vbNewLine & _
		"input.radio {background:" & strPopUpTableColor & ";color:#000000}" & vbNewLine & _
		"-->" & vbNewLine & _
		"</style>" & vbNewLine & _
		"</head>" & vbNewLine & _
		vbNewLine & _
		"<body" & strTmpPageBGImageURL & " bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ link=""" & strLinkColor & """ aLink=""" & strActiveLinkColor & """ vLink=""" & strVisitedLinkColor & """>" & vbNewLine & _
		"<a name=""top""></a><font face=""" & strDefaultFontFace & """>" & vbNewLine & _
		vbNewLine & _
		"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""100%"">" & vbNewLine & _
		"  <tr>" & vbNewLine & _
		"    <td valign=""top"" width=""50%""><a href=""default.asp"" tabindex=""-1"">" & getCurrentIcon(strTitleImage & "||",strForumTitle,"") & "</a></td>" & vbNewLine & _
		"    <td align=""center"" valign=""top"" width=""50%"">" & vbNewLine & _
		"      <table border=""0"" cellPadding=""2"" cellSpacing=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & strForumTitle & "</b></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine
call sForumNavigation()
Response.Write	"</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine

select case Request.Form("Method_Type")

	case "login"
		Response.Write	"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine
		if strLoginStatus = 0 then
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Your username and/or password were incorrect.</font></p>" & vbNewLine & _
					"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Please either try again or register for an account.</font></p>" & vbNewLine
		else
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You logged on successfully!</font></p>" & vbNewLine & _
					"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Thank you for your participation.</font></p>" & vbNewLine
		end if
		Response.Write	"<meta http-equiv=""Refresh"" content=""2; URL=" & Request.ServerVariables("HTTP_REFERER") & """>" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & Request.ServerVariables("HTTP_REFERER") & """>Back To Forum</font></a></p>" & vbNewLine & _
				"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine
		WriteFooter
		Response.End
	case "logout" 
		Response.Write	"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You logged out successfully!</font></p>" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Thank you for your participation.</font></p>" & vbNewLine & _
				"<meta http-equiv=""Refresh"" content=""2; URL=default.asp"">" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""default.asp"">Back To Forum</font></a></p>" & vbNewLine & _
				"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine
		WriteFooter
		Response.End
end select

if (mlev = 0) then
	if not(Instr(Request.ServerVariables("Path_Info"), "register.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "policy.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "pop_profile.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "search.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "login.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "password.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "faq.asp") > 0) then
		Response.Write	"        <form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form1"" name=""form1"">" & vbNewLine & _
				"        <input type=""hidden"" name=""Method_Type"" value=""login"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""center"">" & vbNewLine & _
				"            <table>" & vbNewLine & _
				"              <tr>" & vbNewLine
		if (strAuthType = "db") then
			Response.Write	"                <td><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><b>Username:</b></font><br />" & vbNewLine & _
					"                <input type=""text"" name=""Name"" size=""10"" maxLength=""25"" value=""""></td>" & vbNewLine & _
					"                <td><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><b>Password:</b></font><br />" & vbNewLine & _
					"                <input type=""password"" name=""Password"" size=""10"" maxLength=""25"" value=""""></td>" & vbNewLine & _
					"                <td valign=""bottom"">" & vbNewLine
			if strGfxButtons = "1" then
				Response.Write	"                <input src=""" & strImageUrl & "button_login.gif"" type=""image"" border=""0"" value=""Login"" id=""submit1"" name=""Login"">" & vbNewLine
			else
				Response.Write	"                <input type=""submit"" value=""Login"" id=""submit1"" name=""submit1"">" & vbNewLine
			end if 
			Response.Write	"                </td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td colspan=""3"" align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
					"                <input type=""checkbox"" name=""SavePassWord"" value=""true"" tabindex=""-1"" CHECKED><b> Save Password</b></font></td>" & vbNewLine
		else
			if (strAuthType = "nt") then 
				Response.Write	"                <td><font face=""" & strDefaultFontFace & """ size=""1""  color=""" & strHiLiteFontColor & """>Please <a href=""policy.asp"" tabindex=""-1"">register</a> to post in these Forums</font></td>" & vbNewLine
			end if
		end if 
		Response.Write	"              </tr>" & vbNewLine
		if (lcase(strEmail) = "1") then
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td colspan=""3"" align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
					"                <a href=""password.asp""" & dWStatus("Choose a new password if you have forgotten your current one...") & " tabindex=""-1"">Forgot your "
			if strAuthType = "nt" then Response.Write("Admin ")
			Response.Write	"Password?</a>" & vbNewLine
			if (lcase(strNoCookies) = "1") then
				Response.Write	"                |" & vbNewLine & _
						"                <a href=""admin_home.asp""" & dWStatus("Access the Forum Admin Functions...") & " tabindex=""-1"">Admin Options</a>" & vbNewLine
			end if
			Response.Write	"                <br /><br /></font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
		end if
		Response.Write	"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"        </form>" & vbNewLine
	end if
else
	Response.Write	"        <form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form2"" name=""form2"">" & vbNewLine & _
			"        <input type=""hidden"" name=""Method_Type"" value=""logout"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td align=""center"">" & vbNewLine & _
			"            <table>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>You are logged on as<br />"
	if strAuthType="nt" then
		Response.Write	"<b>" & Session(strCookieURL & "username") & "&nbsp;(" & Session(strCookieURL & "userid") & ")</b></font></td>" & vbNewLine & _
				"                <td>&nbsp;"
	else 
		if strAuthType = "db" then 
			Response.Write	"<b>" & ChkString(strDBNTUserName, "display") & "</b></font></td>" & vbNewLine & _
					"                <td>"
			if strGfxButtons = "1" then
				Response.Write	"<input src=""" & strImageUrl & "button_logout.gif"" type=""image"" border=""0"" value=""Logout"" id=""submit1"" name=""Logout"" tabindex=""-1"">"
			else
				Response.Write	"<input type=""submit"" value=""Logout"" id=""submit1"" name=""submit1"" tabindex=""-1"">"
			end if 
		end if 
	end if 
	Response.Write	"</td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine
	if (mlev = 4) or (lcase(strNoCookies) = "1") then
		Response.Write	"        <tr>" & vbNewLine & _
				"          <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><a href=""admin_home.asp""" & dWStatus("Access the Forum Admin Functions...") & " tabindex=""-1"">Admin Options</a>"
		if mLev = 4 and (strEmailVal = "1" and strRestrictReg = "1" and strEmail = "1" and User_Count > 0) then Response.Write("&nbsp;|&nbsp;<a href=""admin_accounts_pending.asp""" & dWStatus("(" & User_Count & ") Member(s) awaiting approval") & " tabindex=""-1"">(" & User_Count & ") Member(s) awaiting approval</a>")
		Response.Write	"<br /><br /></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine
	end if
	Response.Write	"        </form>" & vbNewLine
end if
Response.Write	"      </table>" & vbNewLine & _
		"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine & _
		"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine
'########### GROUP Categories ########### %>
<!--#INCLUDE FILE="inc_groupjump_to.asp" -->
<% '######## GROUP Categories ##############
Response.Write	"  <tr>" & vbNewLine & _
		"    <td>" & vbNewLine

sub sForumNavigation()
	' DEM --> Added code to show the subscription line
	if strSubscription > 0 and strEmail = "1" then
		if mlev > 0 then
			strSql = "SELECT COUNT(*) AS MySubCount FROM " & strTablePrefix & "SUBSCRIPTIONS"
			strSql = strSql & " WHERE MEMBER_ID = " & MemberID
			set rsCount = my_Conn.Execute (strSql)
			if rsCount.BOF or rsCount.EOF then
				' No Subscriptions found, do nothing
				MySubCount = 0
				rsCount.Close
				set rsCount = nothing
			else
				MySubCount = rsCount("MySubCount")
				rsCount.Close
				set rsCount = nothing
			end if
			if mLev = 4 then
				strSql = "SELECT COUNT(*) AS SubCount FROM " & strTablePrefix & "SUBSCRIPTIONS"
				set rsCount = my_Conn.Execute (strSql)
				if rsCount.BOF or rsCount.EOF then
					' No Subscriptions found, do nothing
					SubCount = 0
					rsCount.Close
					set rsCount = nothing
				else
					SubCount = rsCount("SubCount")
					rsCount.Close
					set rsCount = nothing
				end if
			end if
		else
			SubCount = 0
			MySubCount = 0
		end if
	else
		SubCount = 0
		MySubCount = 0
	end if
	Response.Write	"          <a href=""" & strHomeURL & """" & dWStatus("Homepage") & " tabindex=""-1""><acronym title=""Homepage"">Home</acronym></a>" & vbNewline & _
			"          |" & vbNewline
	if strUseExtendedProfile then 
		Response.Write	"          <a href=""pop_profile.asp?mode=Edit""" & dWStatus("Edit your personal profile...") & " tabindex=""-1""><acronym title=""Edit your personal profile..."">Profile</acronym></a>" & vbNewline
	else
		Response.Write	"          <a href=""javascript:openWindow3('pop_profile.asp?mode=Edit')""" & dWStatus("Edit your personal profile...") & " tabindex=""-1""><acronym title=""Edit your personal profile..."">Profile</acronym></a>" & vbNewline
	end if 
	if strAutoLogon <> "1" then
		if strProhibitNewMembers <> "1" then
			Response.Write	"          |" & vbNewline & _
					"          <a href=""policy.asp""" & dWStatus("Register to post to our forum...") & " tabindex=""-1""><acronym title=""Register to post to our forum..."">Register</acronym></a>" & vbNewline
		end if
	end if
	Response.Write	"          |" & vbNewline & _
			"          <a href=""active.asp""" & dWStatus("See what topics have been active since your last visit...") & " tabindex=""-1""><acronym title=""See what topics have been active since your last visit..."">Active Topics</acronym></a>" & vbNewline 
	' DEM --> Start of code added to show subscriptions if they exist
	if (strSubscription > 0) then
		if mlev = 4 and SubCount > 0 then
			Response.Write	"          |" & vbNewline & _
					"          <a href=""subscription_list.asp?MODE=all""" & dWStatus("See all current subscriptions") & " tabindex=""-1""><acronym title=""See all current subscriptions"">All Subscriptions</acronym></a>" & vbNewline
		end if
		if MySubCount > 0 then
			Response.Write	"          |" & vbNewline & _
					"          <a href=""subscription_list.asp""" & dWStatus("See all of your subscriptions") & " tabindex=""-1""><acronym title=""See all of your subscriptions"">My Subscriptions</acronym></a>" & vbNewline
		end if
	end if
	' DEM --> End of Code added to show subscriptions if they exist
'#######    Begin UserGroup MOD     #######
	blnCanView = ""
	blnCanView = chkUserGroupView(MemberID)
	if blnCanView = true then
		Response.Write	"          |" & vbNewline &_
				"          <a href=""usergroups.asp""" & dWStatus("View UserGroup Information") & " tabindex=""-1""><acronym title=""View UserGroup Information"">UserGroups</acronym></a>" & vbNewline
	end if	
'#######     End UserGroup MOD      #######
	Response.Write	"          |" & vbNewline & _
			"          <a href=""members.asp""" & dWStatus("Current members of these forums...") & " tabindex=""-1""><acronym title=""Current members of these forums..."">Members</acronym></a>" & vbNewline & _
			"          |" & vbNewline & _
			"          <a href=""search.asp"
	if Request.QueryString("FORUM_ID") <> "" then Response.Write("?FORUM_ID=" & cLng(Request.QueryString("FORUM_ID")))
	Response.Write	"""" & dWStatus("Perform a search by keyword, date, and/or name...") & " tabindex=""-1""><acronym title=""Perform a search by keyword, date, and/or name..."">Search</acronym></a>" & vbNewline & _
			"          |" & vbNewline & _
			"          <a href=""faq.asp""" & dWStatus("Answers to Frequently Asked Questions...") & " tabindex=""-1""><acronym title=""Answers to Frequently Asked Questions..."">FAQ</acronym></a>"
end sub

if strGroupCategories = "1" then
	if Session(strCookieURL & "GROUP_NAME") = "" then
		GROUPNAME = " Default Groups "
	else
		GROUPNAME = Session(strCookieURL & "GROUP_NAME")
	end if
	'Forum_SQL - Get Groups
	strSql = "SELECT GROUP_ID, GROUP_CATID " 
	strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
	strSql = strSql & " WHERE GROUP_ID = " & Group
	set rsgroups = Server.CreateObject("ADODB.Recordset")
	rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rsgroups.EOF then
		recGroupCatCount = ""
	else
		allGroupCatData = rsgroups.GetRows(adGetRowsRest)
		recGroupCatCount = UBound(allGroupCatData, 2)
	end if
	rsgroups.Close
	set rsgroups = nothing
end if
%>