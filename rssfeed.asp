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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%

extURL = "" & strForumUrl & "rss.asp"

set xmlDoc = createObject("Msxml.DOMDocument")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
xmlDoc.load(extURL)

If (xmlDoc.parseError.errorCode <> 0) then
	Response.Write	"There was a problem trying to load " & extURL & "<br /><br />"
	Response.Write	"XML error #" & xmlDoc.parseError.errorCode & ": " & xmlDoc.parseError.reason
Else

	set channelNodes = xmlDoc.selectNodes("//channel/*")

	for each entry in channelNodes
		if entry.tagName = "title" then
			strChannelTitle = entry.text
		elseif entry.tagName = "image2" then
			strRSSImage = entry.text
		elseif entry.tagName = "description" then
			strChannelDescription = entry.text
		elseif entry.tagName = "link" then
			strChannelLink = entry.text
		end if
	next

	response.write	"<table class=""content"">" & _
					"<tr class=""header"">" & _
					"<td>" & _
					"<a href=""" & strForumURL & """>" & getCurrentIcon(strTitleImage & "||","","") & "</a>" & _
					"&nbsp;" & strChannelTitle & "&rsquo;s Public RSS Feed" & _
					"</td>" & _
					"</tr>"

	set itemNodes = xmlDoc.selectNodes("//item/*")

	For each item in itemNodes
		Select Case item.tagName
		Case "title"
			strItemTitle = strItemTitle & item.text & "#%#"
		Case "image"
			strRSSImage = strRSSImage & item.text & "#%#"
		Case "link"
			strItemLink = strItemLink & item.text & "#%#"
		Case "description"
			strItemDescription = strItemDescription & item.text & "#%#"
		End Select
	next

	arrItemTitle = split(strItemTitle,"#%#")
	arrItemImage = split(strRSSImage,"#%#")
	arrItemLink = split(strItemLink,"#%#")
	arrItemDescription = split(strItemDescription,"#%#")
	
	for a = 0 to UBound(arrItemTitle) - 1
		response.write	"<tr class=""section"">" & _
						"<td><a href='" & arrItemLink(a) & "'>" & arrItemTitle(a) & "</a></td>" & _
						"</tr>" & _
						"<tr>" & _
						"<td>"
		if strItemDescription <> "" then response.write arrItemDescription(a)
		response.write	"</td>" & _
						"</tr>"
	next
	
	response.write	"</table>"
	
	set channelNodes = nothing
	set itemNodes = nothing

End If

Call WriteFooterShort

%>