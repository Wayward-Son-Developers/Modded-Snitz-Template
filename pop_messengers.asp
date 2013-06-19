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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
	intMemberID = cLng(Request.QueryString("ID"))
else
	intMemberID = 0
end if

select case Request.QueryString("mode")
	case "AIM"
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME, M_AIM "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID = " & intMemberID

		Set rsAIM = my_Conn.execute(strSql)
		
		strProfileName = chkString(rsAIM("M_NAME"),"display")
		strAIM_Name = chkString(rsAIM("M_AIM"),"urlpath")

		Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td>" & strProfileName & "'s AIM Options</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""section"">" & vbNewLine & _
						"<td><b>NOTE:</b> You must have AOL Instant Messenger installed in order for these functions to work properly.</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td><a href=""aim:goIM?screenname=" & strAIM_Name & """ alt=""Opens a send message window to the user."">Send a Message</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td><a href=""aim:goChat?ROOMname=" & strAIM_Name & """>Open a chat room</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td><a href=""aim:addBuddy?screenname=" & strAIM_Name & """>Add to buddy list</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine
		rsAIM.close
		set rsAIM = nothing
	case "ICQ"
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME, M_ICQ "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID = " & intMemberID

		Set rsICQ = my_Conn.execute(strSql)

		strICQ = rsICQ("M_ICQ")
		
		rsICQ.close
		set rsICQ = nothing
		
		Response.Redirect "http://www.icq.com/people/webmsg.php?to=" & strICQ & ""
	case "MSN"
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME, M_MSN "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID = " & intMemberID

		Set rsMSN = my_Conn.execute(strSql)

		strProfileName = chkString(rsMSN("M_NAME"), "display")

		parts = split(rsMSN("M_MSN"),"@")
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

		Response.Write	"<div class=""content""><p>" & strProfileName & "'s MSN Messenger address:</p>" & vbNewLine & _
						"<p><script language=""javascript"" type=""text/javascript"">MSNjs()</script></p></div>" & vbNewLine
		rsMSN.close
		set rsMSN = nothing
end select
if (not(strUseExtendedProfile) and InStr(strReferer, "pop_profile.asp") <> 0) then
	Response.Write	"<p class=""content""><a href=""JavaScript:history.go(-1)"">Return to " & strProfileName & "'s Profile</a></p>" & vbNewLine
end if
WriteFooterShort
Response.End
%>
