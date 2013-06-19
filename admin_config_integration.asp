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
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
%>
<script language="javascript">

function doDefaultHighlights(){
	// Restore highlighted cells to form defaults.
	if (document.Form1.strSiteIntegEnabled[0].defaultChecked) 
		{
		if (document.Form1.strSiteLeft[0].defaultChecked)
			siteleft.style.backgroundColor='#99ff99';
		else
			siteleft.style.backgroundColor='#ff9999';

		if (document.Form1.strSiteHeader[0].defaultChecked)
			siteheader.style.backgroundColor = '#99ff99';
		else
			siteheader.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteRight[0].defaultChecked)
			siteright.style.backgroundColor = '#99ff99';
		else
			siteright.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteFooter[0].defaultChecked)
			sitefooter.style.backgroundColor = '#99ff99';
		else
			sitefooter.style.backgroundColor = '#ff9999';
			
		if (document.Form1.strSiteBorder[0].defaultChecked)
			disptable.border = '1';
		else
			disptable.border = '0';
			
		}
		else {
			siteleft.style.backgroundColor='#cccccc';
			siteheader.style.backgroundColor='#cccccc';
			sitefooter.style.backgroundColor='#cccccc';
			siteright.style.backgroundColor='#cccccc';
			disptable.border = '0';
		}
}

function doHighlights(){
	// Set highlighted cells to current checkbox values.

	if (document.Form1.strSiteIntegEnabled[0].checked) 
		{
		if (document.Form1.strSiteLeft[0].checked)
			siteleft.style.backgroundColor='#99ff99';
		else
			siteleft.style.backgroundColor='#ff9999';

		if (document.Form1.strSiteHeader[0].checked)
			siteheader.style.backgroundColor = '#99ff99';
		else
			siteheader.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteRight[0].checked)
			siteright.style.backgroundColor = '#99ff99';
		else
			siteright.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteFooter[0].checked)
			sitefooter.style.backgroundColor = '#99ff99';
		else
			sitefooter.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteBorder[0].checked)
			disptable.border = '1';
		else
			disptable.border = '0';
			
		}		
		else {
			siteleft.style.backgroundColor='#cccccc';
			siteheader.style.backgroundColor='#cccccc';
			sitefooter.style.backgroundColor='#cccccc';
			siteright.style.backgroundColor='#cccccc';
			disptable.border = '0'
		}
}
</script>
<%


Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Site&nbsp;Integration&nbsp;Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""

	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLstring"))
			end if
		next

		Application(strCookieURL & "ConfigLoaded") = ""

		Call OkMessage("Configuration Posted!","admin_home.asp","Back To Admin Home")
	else
		Call FailMessage(Err_Msg,True)
	end if
else
	Response.Write	"<form action=""admin_config_integration.asp"" method=""post"" id=""Form1"" name=""Form1"" onReset=""Javascript:doDefaultHighlights();"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Site Integration Configuration</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Site Integration features:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strSiteIntegEnabled"" value=""1""" & chkRadio(strSiteIntegEnabled,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strSiteIntegEnabled"" value=""0""" & chkRadio(strSiteIntegEnabled,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Version info:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">[<b>v1.2</b>]</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td width=""400"" colspan=""2"">" & vbNewLine & _
					"<table id=""disptable"" width=""380"" class=""contentcontainer"" style=""border:"
	if strSiteBorder = "1" and strSiteIntegEnabled = "1" then
		response.write "2"
	else
		response.write "1"
	end if
	response.write	"px solid black;"">" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td name=""siteheader"" id=""siteheader"" colspan=""3"" style=""background-color:"
	if strSiteIntegEnabled = "0" then
		response.write "#cccccc"
	else
		if strSiteHeader = "1" then
			response.write "#99ff99"
		else
			response.write "#ff9999"
		end if
	end if
	response.write  ";text-align:center;"">Site Header</td></tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td id=""siteleft"" width=""40"" style=""background-color:"
	if strSiteIntegEnabled = "0" then
		response.write "#cccccc"
	else
		if strSiteLeft = "1" then
			response.write "#99ff99"
		else
			response.write "#ff9999;"
		end if
	end if
	response.write  ";text-align:center;"">Site<br />Left</td>" & vbNewLine & _
					"<td width=""300"" style=""background:white;text-align:center;""><img src=""" & strImageURL & "logo_snitz_forums_2000.gif"" alt=""Forums"" width=""163"" height=""76"" border=""0""></td>" & vbNewLine & _
					"<td id=""siteright"" width=""40"" style=""background-color:"
	if strSiteIntegEnabled = "0" then
		response.write "#cccccc"
	else
		if strSiteRight = "1" then
			response.write "#99ff99"
		else
			response.write "#ff9999"
		end if
	end if
	response.write  ";text-align:center;"">Site<br />Right</td></tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td id=""sitefooter"" colspan=""3"" style=""background-color:"
	if strSiteIntegEnabled = "0" then
		response.write "#cccccc"
	else
		if strSiteFooter = "1" then
			response.write "#99ff99"
		else
			response.write "#ff9999"
		end if
	end if
	response.write	";text-align:center;"">Site Footer</td></tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Site Header:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strSiteHeader"" value=""1""" & chkRadio(strSiteHeader,0,false) & " onClick=""javascript:doHighlights();"">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strSiteHeader"" value=""0""" & chkRadio(strSiteHeader,0,true) & " onClick=""javascript:doHighlights();"">" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Site Left:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strSiteLeft"" value=""1""" & chkRadio(strSiteLeft,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strSiteLeft"" value=""0""" & chkRadio(strSiteLeft,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Site Right:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strSiteRight"" value=""1""" & chkRadio(strSiteRight,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strSiteRight"" value=""0""" & chkRadio(strSiteRight,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Site Footer:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strSiteFooter"" value=""1""" & chkRadio(strSiteFooter,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strSiteFooter"" value=""0""" & chkRadio(strSiteFooter,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Show Borders:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"On: <input type=""radio"" class=""radio"" name=""strSiteBorder"" value=""1""" & chkRadio(strSiteBorder,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"Off: <input type=""radio"" class=""radio"" name=""strSiteBorder"" value=""0""" & chkRadio(strSiteBorder,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
					"</td>" & vbNewLine & _
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
