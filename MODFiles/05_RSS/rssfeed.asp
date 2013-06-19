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
<!--#INCLUDE FILE="inc_func_common.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<%

extURL = "" & strForumUrl & "rss.asp"

set xmlDoc = createObject("Msxml.DOMDocument")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
xmlDoc.load(extURL)

If (xmlDoc.parseError.errorCode <> 0) then
	Response.Write "XML error: " & xmlDoc.parseError.reason
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

	Response.Write	"<style type=""text/css"">" & vbNewLine & _
		"<!--" & vbNewLine & _
		"/* ##### Extended Color Code Mod ##### */"  & vbNewLine & _
		"body {Scrollbar-Face-Color:" & strScrollbarFaceColor & ";Scrollbar-Arrow-Color:" & strScrollbarArrowColor & ";Scrollbar-Track-Color:" & strScrollbarTrackColor & ";Scrollbar-Shadow-Color:" & strScrollbarShadowColor & ";Scrollbar-Highlight-Color:" & strScrollbarHighlightColor & ";Scrollbar-3Dlight-Color:" & strScrollbar3DlightColor & "}" & vbNewLine & _
		"a:link    {color:" & strLinkColor & ";background-color:" & strLinkBGColor & ";text-decoration:" & strLinkTextDecoration & "}" & vbNewLine & _
		"a:visited {color:" & strVisitedLinkColor & ";background-color:" & strVisitedLinkBGColor & ";text-decoration:" & strVisitedTextDecoration & "}" & vbNewLine & _
		"a:hover   {color:" & strHoverFontColor & ";background-color:" & strHoverFontBGColor & ";text-decoration:" & strHoverTextDecoration & "}" & vbNewLine & _
		"a:active  {color:" & strActiveLinkColor & ";background-color:" & strActiveLinkBGColor & ";text-decoration:" & strActiveTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:link    {color:" & strForumLinkColor & ";background-color:" & strForumLinkBGColor & ";text-decoration:" & strForumLinkTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:visited {color:" & strForumVisitedLinkColor & ";background-color:" & strForumVisitedLinkBGColor & ";text-decoration:" & strForumVisitedTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:hover   {color:" & strForumHoverFontColor & ";background-color:" & strForumHoverFontBGColor & ";text-decoration:" & strForumHoverTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:active  {color:" & strForumActiveLinkColor & ";background-color:" & strForumActiveLinkBGColor & ";text-decoration:" & strForumActiveTextDecoration & "}" & vbNewLine & _
		".spnSearchHighlight {background-color:" & strSearchHiLiteColor & "}" & vbNewLine & _
		"select {background-color:" & StrForumCellColor & "; color: " & StrDefaultFontColor & "; border-width:0; border-color:" & StrTableborderColor & "}" & vbNewLine & _
		"textarea {background-color:" & StrForumCellColor & "; color: " & StrDefaultFontColor & "; border-width:1; border-color:""#000000""}" & vbNewLine & _
		"input.buttons {background-color:" & StrForumCellColor & "; color: " & StrDefaultFontColor & "; border-width:1; border-color:" & StrTableborderColor & "}" & vbNewLine & _
		"input.buttons2 {background-color:" & StrAltForumCellColor & "; color: " & StrDefaultFontColor & "; border-width:1; border-color:" & StrTableborderColor & "}" & vbNewLine & _
		"input.buttons3 {background-color:" & StrPopupTableColor & "; color: " & StrForumFontColor & "; border-width:2; border-color:" & StrPopupborderColor & "}" & vbNewLine & _
		"input.newLogin {background-color:" & StrForumCellColor & "; color:" & StrDefaultFontColor & "; border-width:1; border-color:""#000000""}" & vbNewLine & _
		"input.search {background-color:" & StrAltForumCellColor & "; color:" & StrDefaultFontColor & "; border-width:1; border-color:""#000000""}" & vbNewLine & _
		"input.radio {background-color:""; color:#000000}" & vbNewLine & _
		"-->" & vbNewLine & _
		"</style>" & vbNewLine

	response.write "<body class=""pb dfs dff dfc"">"
	response.write "<table border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & _
					"<tr><td bgcolor=""" & strHeadCellColor & """>"
	response.write "<a href=""http://www.eastoverfd.com"">" & getCurrentIcon(strTitleImage & "||","","") & "</a>"
	response.write "</td><td width=""100%"" bgcolor=""" & strHeadCellColor & """>"
	'response.write "<font face=""Tahoma"" size=""5"" color=""" & strHeadFontColor & """>&nbsp;" & strChannelTitle & " RSS Feed</font>"
	response.write "</td></tr></table><br />"

	set itemNodes = xmlDoc.selectNodes("//item/*")

	For each item in itemNodes
		if item.tagName = "title" then
			strItemTitle = strItemTitle & item.text & "#%#"
		elseif item.tagName = "image" then
			strRSSImage = strRSSImage & item.text & "#%#"
		elseif item.tagName = "link" then
			strItemLink = strItemLink & item.text & "#%#"
		elseif item.tagName = "description" then
			strItemDescription = strItemDescription & item.text & "#%#"
		end if
	next

	arrItemTitle = split(strItemTitle,"#%#")
	arrItemImage = split(strRSSImage,"#%#")
	arrItemLink = split(strItemLink,"#%#")
	arrItemDescription = split(strItemDescription,"#%#")

	response.write ""
		for a = 0 to UBound(arrItemTitle) - 1
			response.write "<table border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""2"" bgcolor=""" & strPopupTableColor & """><tr><td bgcolor=""" & strHeadCellColor & """>"
			response.write "<b><font face=""Tahoma"" size=""4""><a href='" & arrItemLink(a) & "'>" & arrItemTitle(a) & "</a></font></b>"
			response.write "</td></tr><tr><td bgcolor=""" & strForumCellColor & """>"
				if strItemDescription <> "" then
					response.write "<font face=""Tahoma"" size=""2"" color=""" & strforumFontColor & """>" & arrItemDescription(a)
				end if
			response.write "</font></td></tr></table><br />"
		next
	response.write ""

	set channelNodes = nothing
	set itemNodes = nothing

End If

%>