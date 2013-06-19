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
Response.Write	"<table class=""misc"" width=""100%"">" & vbNewLine & _
		"<tr>" & vbNewLine & _
		"<td class=""secondnav"">" & vbNewLine & _
		getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
		getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Font/Table&nbsp;Color&nbsp;Code&nbsp;Configuration</td>" & vbNewLine & _
		"</tr>" & vbNewLine & _
		"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strTopicWidthLeft") = "" then
		Err_Msg = Err_Msg & "<li>You Must enter a value for the Topic Left Column Width</li>"
	end if
	if Request.Form("strTopicWidthRight") = "" then
		Err_Msg = Err_Msg & "<li>You Must enter a value for the Topic Right Column Width</li>"
	end if

	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next

		Application(strCookieURL & "ConfigLoaded") = ""

		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Configuration Posted!</font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""admin_home.asp"">Back To Admin Home</font></a></p>" & vbNewLine
	else
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
else
	Response.Write	"    <form action=""admin_config_colors.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"    <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
			"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Font/Table Color Code Configuration</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Font Face Type:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strDefaultFontFace"" size=""25"" maxLength=""250"" value=""" & chkExist(strDefaultFontFace) & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontfacetype')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Default Font Size:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strDefaultFontSize"">" & vbNewLine & _
			"                	<option value=""""" & chkSelect(strDefaultFontSize,"") & ">None (blank)</option>" & vbNewLine & _
			"                	<option value=""1""" & chkSelect(strDefaultFontSize,1) & ">1 (8 pt)</option>" & vbNewLine & _
			"                	<option value=""2""" & chkSelect(strDefaultFontSize,2) & ">2 (10 pt)</option>" & vbNewLine & _
			"                	<option value=""3""" & chkSelect(strDefaultFontSize,3) & ">3 (12 pt)</option>" & vbNewLine & _
			"                	<option value=""4""" & chkSelect(strDefaultFontSize,4) & ">4 (14 pt)</option>" & vbNewLine & _
			"                	<option value=""5""" & chkSelect(strDefaultFontSize,5) & ">5 (18 pt)</option>" & vbNewLine & _
			"                	<option value=""6""" & chkSelect(strDefaultFontSize,6) & ">6 (24 pt)</option>" & vbNewLine & _
			"                	<option value=""7""" & chkSelect(strDefaultFontSize,7) & ">7 (36 pt)</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontsize')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Header Font Size:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strHeaderFontSize"">" & vbNewLine & _
			"                	<option value=""""" & chkSelect(strHeaderFontSize,"") & ">None (blank)</option>" & vbNewLine & _
			"                	<option value=""1""" & chkSelect(strHeaderFontSize,1) & ">1 (8 pt)</option>" & vbNewLine & _
			"                	<option value=""2""" & chkSelect(strHeaderFontSize,2) & ">2 (10 pt)</option>" & vbNewLine & _
			"                	<option value=""3""" & chkSelect(strHeaderFontSize,3) & ">3 (12 pt)</option>" & vbNewLine & _
			"                	<option value=""4""" & chkSelect(strHeaderFontSize,4) & ">4 (14 pt)</option>" & vbNewLine & _
			"                	<option value=""5""" & chkSelect(strHeaderFontSize,5) & ">5 (18 pt)</option>" & vbNewLine & _
			"                	<option value=""6""" & chkSelect(strHeaderFontSize,6) & ">6 (24 pt)</option>" & vbNewLine & _
			"                	<option value=""7""" & chkSelect(strHeaderFontSize,7) & ">7 (36 pt)</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontsize')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Footer Font Size:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strFooterFontSize"">" & vbNewLine & _
			"                	<option value=""""" & chkSelect(strFooterFontSize,"") & ">None (blank)</option>" & vbNewLine & _
			"                	<option value=""1""" & chkSelect(strFooterFontSize,1) & ">1 (8 pt)</option>" & vbNewLine & _
			"                	<option value=""2""" & chkSelect(strFooterFontSize,2) & ">2 (10 pt)</option>" & vbNewLine & _
			"                	<option value=""3""" & chkSelect(strFooterFontSize,3) & ">3 (12 pt)</option>" & vbNewLine & _
			"                	<option value=""4""" & chkSelect(strFooterFontSize,4) & ">4 (14 pt)</option>" & vbNewLine & _
			"                	<option value=""5""" & chkSelect(strFooterFontSize,5) & ">5 (18 pt)</option>" & vbNewLine & _
			"                	<option value=""6""" & chkSelect(strFooterFontSize,6) & ">6 (24 pt)</option>" & vbNewLine & _
			"                	<option value=""7""" & chkSelect(strFooterFontSize,7) & ">7 (36 pt)</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontsize')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Base Background Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strPageBGColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strPageBGColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Default Font Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strDefaultFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strDefaultFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strLinkColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strLinkTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strLinkTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strLinkTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strLinkTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strLinkTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strLinkTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Visited Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strVisitedLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strVisitedLinkColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Visited Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strVisitedTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strVisitedTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strVisitedTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strVisitedTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strVisitedTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strVisitedTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>" & vbNewLine & _
			"                </td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Active Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strActiveLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strActiveLinkColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Active Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strActiveTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strActiveTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strActiveTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strActiveTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strActiveTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strActiveTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>" & vbNewLine & _
			"                </td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Hover Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strHoverFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHoverFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Hover Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strHoverTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strHoverTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strHoverTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strHoverTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strHoverTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strHoverTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Header Background Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strHeadCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHeadCellColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Header Font Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strHeadFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHeadFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Category Background Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strCategoryCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strCategoryCellColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Category Font Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strCategoryFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strCategoryFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>First Cell Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumFirstCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumFirstCellColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>First Alternating Cell Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumCellColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Second Alternating Cell Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strAltForumCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strAltForumCellColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Font Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumLinkColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strForumLinkTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strForumLinkTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumLinkTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumLinkTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumLinkTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumLinkTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Visited Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumVisitedLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumVisitedLinkColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Visited Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strForumVisitedTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strForumVisitedTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumVisitedTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumVisitedTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumVisitedTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumVisitedTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>" & vbNewLine & _
			"                </td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Active Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumActiveLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumActiveLinkColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Active Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strForumActiveTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strForumActiveTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumActiveTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumActiveTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumActiveTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumActiveTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>" & vbNewLine & _
			"                </td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Hover Link Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strForumHoverFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumHoverFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>In Forum Hover Link Decoration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strForumHoverTextDecoration"">" & vbNewLine & _
			"                	<option" & chkSelect(strForumHoverTextDecoration,"none") & ">none</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumHoverTextDecoration,"blink") & ">blink</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumHoverTextDecoration,"line-through") & ">line-through</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumHoverTextDecoration,"overline") & ">overline</option>" & vbNewLine & _
			"                	<option" & chkSelect(strForumHoverTextDecoration,"underline") & ">underline</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Table Border Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strTableBorderColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strTableBorderColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Pop-Up Table Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strPopUpTableColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strPopUpTableColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Pop-Up Table Border Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strPopUpBorderColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strPopUpBorderColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>New Font Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strNewFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strNewFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>HighLight Font Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strHiLiteFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHiLiteFontColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Search HighLight Color:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strSearchHiLiteColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strSearchHiLiteColor) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Page Background Image URL:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strPageBGImageURL"" size=""25"" maxLength=""100"" value=""" & chkExist(strPageBGImageURL) & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#pagebgimage')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Table Size Configuration</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Topic Left Column Width:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strTopicWidthLeft"" size=""5"" maxLength=""4"" value=""" & chkExistElse(strTopicWidthLeft,"100") & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#columnwidth')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Topic NOWRAP Left:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strTopicNoWrapLeft"" value=""1""" & chkRadio(strTopicNoWrapLeft,0,false) & ">" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strTopicNoWrapLeft"" value=""0""" & chkRadio(strTopicNoWrapLeft,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#nowrap')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Topic Right Column Width:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""strTopicWidthRight"" size=""5"" maxLength=""4"" value=""" & chkExistElse(strTopicWidthRight,"100%") & """>" & vbNewLine & _
			"	         <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#columnwidth')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Topic NOWRAP Right:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strTopicNoWrapRight"" value=""1""" & chkRadio(strTopicNoWrapRight,0,false) & ">" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strTopicNoWrapRight"" value=""0""" & chkRadio(strTopicNoWrapRight,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#nowrap')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"    </form>" & vbNewLine
end if
WriteFooter
Response.End
%>
