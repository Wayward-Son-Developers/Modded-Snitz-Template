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
%><!--#INCLUDE FILE="inc_func_common.asp" --><%

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
strReferer = chkString(request.servervariables("HTTP_REFERER"),"refer")

if Application(strCookieURL & "down") then 
	if not Instr(strScriptName,"admin_") > 0 And not Instr(strScriptName,"down.asp") > 0 Then
		Response.redirect("down.asp")
	end if
end if

If strDBType = "" then 
	Response.Write	"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbNewLine & _
					"<html>" & vbNewLine & _
					"<head>" & vbNewline & _
					"<title>" & strForumTitle & "</title>" & vbNewline & _
					"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" />" & vbNewLine
	
	'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
	'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	
	Response.Write	"<link rel=""stylesheet"" type=""text/css"" href=""inc_style_main.css"" />"
	Response.Write	"</head>" & vbNewLine & _
					"<body>" & vbNewLine & _
					"<table class=""oops"" cellspacing=""0"" width=""50%"" height=""40%"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td><p><b>There has been a problem...</b><br /><br />" & _
					"Your <b>strDBType</b> is not set, please edit your <b>config.asp</b> to reflect your database type.</p>" & vbNewLine & _
					"<p><a href=""default.asp"" target=""_top"">Click here to retry.</a></p></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
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
			Group = cLng(Request.Cookies(strCookieURL & "GROUP"))
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
	if not Instr(strScriptName,"register.asp") > 0 and _
	not Instr(strScriptName,"password.asp") > 0 and _
	not Instr(strScriptName,"faq.asp") > 0 and _
	not Instr(strScriptName,"contact.asp") > 0 and _
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

strRssURL = BuildRssURL(MemberID)

if Session(strCookieURL & "UserGroups" & MemberID) = "" or IsNull(Session(strCookieURL & "UserGroups" & MemberID)) then
	strGroupMembership = getGroupMembership(MemberID,1)
	Session(strCookieURL & "UserGroups" & MemberID) = strGroupMembership
	Session(strCookieURL & "UserGroups" & MemberID) = strGroupMembership
end if

Response.Write	"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbNewLine & _
				"<html>" & vbNewline & vbNewline & _
				"<head>" & vbNewline & _
				"<title>" & GetNewTitle(strScriptName) & "</title>" & vbNewline & _
				"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" />" & vbNewLine

'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

Response.Write	"<link rel=""alternate"" type=""application/rss+xml"" href=""" & strForumUrl & "rss.asp"" title=""" & strForumTitle & "'s Public RSS Feed"">" & vbNewLine
Response.Write	"<link rel=""stylesheet"" type=""text/css"" href=""inc_style_main.css"" />" & vbNewLine
Response.Write	"<script type=""text/javascript"" src=""inc_window.js""></script>" & vbNewLine
Response.Write	"<link href=""" & strImageUrl & "favicon.ico"" rel=""icon"" type=""image/ico"" />" & vbNewLine
Response.Write	"</head>" & vbNewLine & _
				"<body>" & vbNewLine & _
				"<a name=""top""></a>" & vbNewLine
if strSiteIntegEnabled = "1" then
	Response.Write "<table width=""100%"" border="""
	if strSiteBorder = "1" then
		Response.Write "1"
	else
		Response.Write "0"
	end if
	Response.Write """ cellspacing=""0"">" & vbNewLine
	if strSiteHeader = "1" then
		Response.Write	"<tr>" & vbNewLine & "<td"
		if strSiteLeft = "1" or strSiteRight = "1" then
			if strSiteLeft = "1" and strSiteRight = "1" then
				Response.Write " colspan=""3"""
			else
				Response.Write " colspan=""2"""
			end if
		end if
		Response.Write	">"
		%><!--#include file="inc_site_header.asp"--><%
		Response.Write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine
	end if
	Response.Write	"<tr>" & vbNewLine & _
					"<td valign=""top"">" & vbNewLine
	if strSiteLeft = "1" then
		%><!--#include file="inc_site_left.asp"--><%
		Response.Write	"</td>" & vbNewLine & _
						"<td valign=""top"">" & vbNewLine
	end if
end if
Response.Write	"<table class=""masthead"" width=""95%"" cellSpacing=""0"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td><div class=""logo""><a href=""default.asp"" tabindex=""-1"">" & getCurrentIcon(strTitleImage & "||",strForumTitle," class=""logo""") & "</a></div><br />" & vbNewLine & _
				"<div class=""title"">" & strForumTitle & "</div><br /><div class=""navtxt"">" & vbNewLine
				Call sForumNavigation()
Response.Write	"</div></td>" & vbNewLine & _
				"</tr>" & vbNewLine

if (mlev = 0) then
	if not(Instr(Request.ServerVariables("Path_Info"), "register.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "pop_profile.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "search.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "login.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "password.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "faq.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "contact.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "down.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "post.asp") > 0) then
		Response.Write	"<form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form1"" name=""form1"">" & vbNewLine & _
						"<input type=""hidden"" name=""Method_Type"" value=""login"">" & vbNewLine & _
						"<tr>" & vbNewLine & "<td>" & vbNewLine
		if (strAuthType = "db") then
			Response.Write	"<b>Username:</b>&nbsp;" & vbNewLine & _
							"<input type=""text"" name=""Name"" size=""10"" maxLength=""25"" value="""">&nbsp;" & vbNewLine & _
							"<b>Password:</b>&nbsp;" & vbNewLine & _
							"<input type=""password"" name=""Password"" size=""10"" maxLength=""25"" value="""">&nbsp;" & vbNewLine
			if strGfxButtons = "1" then
				Response.Write	"<input class=""gfxbutton"" src=""" & strImageUrl & "button_login.gif"" type=""image"" border=""0"" value=""Login"" id=""submit1"" name=""Login"">" & vbNewLine
			else
				Response.Write	"<button type=""submit"" id=""submit1"" name=""submit1"">Login</button>" & vbNewLine
			end if 
			Response.Write	"<br /><input type=""checkbox"" name=""SavePassWord"" value=""true"" tabindex=""-1"" CHECKED> <b>Save Password</b>" & vbNewLine
		else
			if (strAuthType = "nt") then 
				Response.Write	"Please <a href=""register.asp"" tabindex=""-1"">register</a> to post in these Forums." & vbNewLine
			end if
		end if 
		if (lcase(strEmail) = "1") then
			Response.Write	"&nbsp;(<a href=""password.asp""" & dWStatus("Choose a new password if you have forgotten your current one...") & " tabindex=""-1"">Forgot your "
			if strAuthType = "nt" then Response.Write("Admin ")
			Response.Write	"Password?</a>)" & vbNewLine
		end if
		Response.Write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine
		Response.Write	"</form>" & vbNewLine
	end if
else
	Response.Write	"<form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form2"" name=""form2"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""logout"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>You are logged on as:&nbsp;"
	if strAuthType="nt" then
		Response.Write	"<b>" & Session(strCookieURL & "username") & "(" & Session(strCookieURL & "userid") & ")</b>" & vbNewLine
	else 
		if strAuthType = "db" then 
			Response.Write	"<b>" & profileLink(ChkString(strDBNTUserName, "display"),MemberID) & "</b>&nbsp;" & vbNewLine
			if strGfxButtons = "1" then
				Response.Write	"<input class=""gfxbutton"" src=""" & strImageUrl & "button_logout.gif"" type=""image"" border=""0"" value=""Logout"" id=""submit1"" name=""Logout"" tabindex=""-1"">"
			else
				Response.Write	"<button type=""submit"" id=""submit1"" name=""submit1"" tabindex=""-1"">Logout</button>"
			end if 
		end if 
	end if 
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	Response.Write	"</form>" & vbNewLine
end if
Response.Write "</table>" & vbNewLine

'Login/Logout message
select case Request.Form("Method_Type")
	case "login"
		if strLoginStatus = 0 then
			Call FailMessage("<li>Your username and/or password were incorrect. Please either try again or register for an account.</li>",True)
		else
			Call OkMessage("<p>You logged on successfully!</p>",strReferer,"Back To Forum")
		end if
		WriteFooter
		Response.End
	case "logout" 
		Call OkMessage("<p>You logged out successfully!</p>","default.asp","Back To Forum")
		WriteFooter
		Response.End
end select

'Start Main page Content
Response.Write "<table class=""contentcontainer"" width=""95%"" cellSpacing=""0"">" & vbNewLine
'########### GROUP Categories ########### %>
<!--#INCLUDE FILE="inc_groupjump_to.asp" -->
<% '######## GROUP Categories ##############
Response.Write	"<tr>" & vbNewLine & _
				"<td>" & vbNewLine

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
	Response.Write	"<a href=""" & strHomeURL & """" & dWStatus("Homepage") & " tabindex=""-1"">Home</a>" & vbNewline & _
					" | " & vbNewline
	if strUseExtendedProfile then 
		Response.Write	"<a href=""pop_profile.asp?mode=Edit""" & dWStatus("Edit your personal profile...") & " tabindex=""-1"">Profile</a>" & vbNewline
	else
		Response.Write	"<a href=""javascript:openWindow3('pop_profile.asp?mode=Edit')""" & dWStatus("Edit your personal profile...") & " tabindex=""-1"">Profile</a>" & vbNewline
	end if 
	if strAutoLogon <> "1" then
		if strProhibitNewMembers <> "1" then
			Response.Write	" | <a href=""register.asp""" & dWStatus("Register to post to our forum...") & " tabindex=""-1"">Register</a>" & vbNewline
		end if
	end if
	Response.Write	" | <a href=""active.asp""" & dWStatus("See what topics have been active since your last visit...") & " tabindex=""-1"">Active Topics</a>" & vbNewline 
	' DEM --> Start of code added to show subscriptions if they exist
	if (strSubscription > 0) then
		if mlev = 4 and SubCount > 0 then Response.Write	" | <a href=""subscription_list.asp?MODE=all""" & dWStatus("See all current subscriptions") & " tabindex=""-1"">All Subscriptions</a>" & vbNewline
		if MySubCount > 0 then Response.Write	" | <a href=""subscription_list.asp""" & dWStatus("See all of your subscriptions") & " tabindex=""-1"">My Subscriptions</a>" & vbNewline
	end if
	' DEM --> End of Code added to show subscriptions if they exist
	if chkUserGroupView(MemberID) = true then Response.Write    " | <a href=""usergroups.asp""" & dWStatus("View UserGroup Information") & " tabindex=""-1"">UserGroups</a>" & vbNewline
	Response.Write	" | <a href=""members.asp""" & dWStatus("Current members of these forums...") & " tabindex=""-1"">Members</a>" & vbNewline & _
					" | <a href=""search.asp"
	if Request.QueryString("FORUM_ID") <> "" then Response.Write("?FORUM_ID=" & cLng(Request.QueryString("FORUM_ID")))
	Response.Write	"""" & dWStatus("Perform a search by keyword, date, and/or name...") & " tabindex=""-1"">Search</a>" & vbNewline & _
					" | <a href=""faq.asp""" & dWStatus("Answers to Frequently Asked Questions...") & " tabindex=""-1"">FAQ</a>" & vbNewLine
	If strEmail = "1" Then Response.Write " | <a href=""contact.asp""" & dWStatus("Contact Us") & " tabindex=""-1"">Contact</a>" & vbNewline
	if (mlev = 4) or (lcase(strNoCookies) = "1") then Response.Write " | <a href=""admin_home.asp""" & dWStatus("Access the Forum Admin Functions...") & " tabindex=""-1"">Admin Options</a>"
	if mLev = 4 and (strEmailVal = "1" and strRestrictReg = "1" and strEmail = "1" and User_Count > 0) then Response.Write(" | <a href=""admin_accounts_pending.asp""" & dWStatus("(" & User_Count & ") Member(s) awaiting approval") & " tabindex=""-1"">(" & User_Count & ") Member(s) awaiting approval</a>")
end sub
%>
