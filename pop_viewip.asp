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
<!--#include file="config.asp"-->
<% 
Forum_ID = cLng("0" & request.querystring("FORUM_ID"))
Topic_ID = cLng("0" & request.querystring("TOPIC_ID"))
Reply_ID = cLng("0" & request.querystring("REPLY_ID"))
select case Request.QueryString("mode")
	case "getIP"
		if request("ARCHIVE") = "true" then
			strActivePrefix = strTablePrefix & "A_"
			ArchiveView = "true"
			ArchiveLink = "ARCHIVE=true&"
		else
			strActivePrefix = strTablePrefix
			ArchiveView = ""
			ArchiveLink = ""
		end if
%>
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
		DisplayIP
		IP = ""
		Title = ""

		WriteFooterShort
		Response.End
	case else
		'## Do Nothing - Continue
end select

sub DisplayIP()
	usr = (chkForumModerator(Forum_ID , strDBNTUserName))
	if (chkUser((strDBNTUserName), (Request.Cookies(strUniqueID & "User")("Pword")), -1) = 4) then 
		usr = 1
	end if
	if usr = 1 then
		if Topic_ID <> 0 then
			'## Forum_SQL
			strSql = "SELECT T_IP "
			strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			set rsIP = my_Conn.Execute(strSql)

			IP = rsIP("T_IP")
		else
			if Reply_ID <> 0 then
				'## Forum_SQL
				strSql = "SELECT R_IP "
				strSql = strSql & " FROM " & strActivePrefix & "REPLY "
				strSql = strSql & " WHERE REPLY_ID = " & Reply_ID

				set rsIP = my_Conn.Execute(strSql)

				IP = rsIP("R_IP")
			end if
		end if
		set rsIP = nothing
		Call PopOkMessage("<b>User's IP address:</b><br /><a href=""" & strIPLookup & ip & """ target=""_blank"">" & ip & "</a>",False)
	else
		Call FailMessage("<li>Only moderators and administrators can perform this action.</li>",False)
	end If
end sub
%>
