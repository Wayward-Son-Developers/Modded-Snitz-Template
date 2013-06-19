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

Response.Write	"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine
' ^ End Main page Content

' Start Footer Content
Response.Write	"<table class=""footer"" width=""95%"" cellspacing=""0"">" & vbNewLine

Response.Write	"<tr>" & vbNewLine & _
				"<td class=""footertitle"">" & strForumTitle & "</td>" & vbNewLine
If strShowTimer = "1" Then
	Response.Write	"<td class=""timertext"">" & chkString(replace(strTimerPhrase, "[TIMER]", abs(round(StopTimer(1), 2)), 1, -1, 1),"display") & "</td>" & vbNewLine
End If
Response.Write	"<td class=""copyright"">&copy; " & strCopyright & "&nbsp;" & vbNewLine & _
				"<a href=""#top""" & dWStatus("Go To Top Of Page...") & " tabindex=""-1"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
				"</tr>" & vbNewLine

Response.Write	"<tr class=""footernav""><td colspan=""" & (2+strShowTimer) & """>" & vbNewLine
				Call sForumNavigation()
Response.Write	"</td></tr>" & vbNewLine

Response.Write	"<tr><td style=""text-align:center;"" colspan=""" & (2+strShowTimer) & """>"
Response.Write	"<a href=""rss.asp"" target=""_blank"">" & getCurrentIcon(strIconRSSPublicFeed,"Public RSS Feed"," border=""0""") & "</a>&nbsp;&nbsp;" & _
				"<a href=""rssfeed.asp"" target=""_blank"">" & getCurrentIcon(strIconRSSPublicReader,"Web Reader for Public RSS Feed"," border=""0""") & "</a>"
If mLev > 0 Then Response.Write("&nbsp;&nbsp;<a href=""" & strRssURL & """ target=""_blank"">" & getCurrentIcon(strIconRSSPrivateFeed,"Personal RSS Feed"," border=""0""") & "</a>")
Response.Write	"</td></tr>"


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<tr><td style=""text-align:center;"" colspan=""" & (2+strShowTimer) & """>"
Response.Write	"<a href=""http://forum.snitz.com"" target=""_blank"" tabindex=""-1""><acronym title=""Powered By: " & strVersion & """>"
if strShowImagePoweredBy = "1" then 
	Response.Write	getCurrentIcon("logo_powered_by.gif||","Powered By: " & strVersion,"")
else
	Response.Write	"Powered By: Snitz Forums"
end if
Response.Write	"</acronym></a></td></tr>" & vbNewLine
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

Response.Write	"</table>" & vbNewLine
' ^ End Footer Content

if strSiteIntegEnabled = "1" then
	if strSiteRight = "1" then
		Response.Write	"</td>" & vbNewLine & _
						"<td valign=""top"">" & vbNewLine
		%><!--#include file="inc_site_right.asp"--><%
	end if
	Response.Write	"</td>" & vbNewLine & _
	"</tr>" & vbNewLine
	if strSiteFooter = "1" then
		Response.write  "<tr>" & vbNewLine & "<td"
		if strSiteLeft = "1" or strSiteRight = "1" then
			if strSiteLeft = "1" and strSiteRight = "1" then
				Response.write " colspan=""3"""
			else
				Response.write " colspan=""2"""
			end if
		end if
		Response.write	">"
		%><!--#include file="inc_site_footer.asp"--><%
		Response.write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine
	end if
	Response.write "</table>" & vbNewLine
end if

response.write "</body>" & vbNewLine & _
"</html>" & vbNewLine

my_Conn.Close
set my_Conn = nothing 
%>
