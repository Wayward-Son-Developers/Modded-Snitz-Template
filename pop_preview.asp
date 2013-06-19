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
<!--#INCLUDE FILE="inc_header_short.asp"-->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<!--#INCLUDE FILE="inc_func_secure.asp"-->
<%
Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
				"      function submitPreview()" & vbNewLine & _
				"      {" & vbNewLine & _
				"      		if (window.opener.document.PostTopic.Subject) {" & vbNewLine & _
				"      			document.previewTopic.subject.value = window.opener.document.PostTopic.Subject.value;" & vbNewLine & _
				"      		}" & vbNewLine & _
				"      		document.previewTopic.message.value = window.opener.document.PostTopic.Message.value;" & vbNewLine & _
				"      		if (window.opener.document.PostTopic.Sig) {" & vbNewLine & _
				"      			if (window.opener.document.PostTopic.Sig.checked) {" & vbNewLine & _
				"      				document.previewTopic.sig.value = ""yes""" & vbNewLine & _
				"      			}" & vbNewLine & _
				"      		}" & vbNewLine & _
				"      		if (window.opener.document.PostTopic.Author) {" & vbNewLine & _
				"      			document.previewTopic.author.value = window.opener.document.PostTopic.Author.value;" & vbNewLine & _
				"      		}" & vbNewLine & _
				"      		document.previewTopic.submit()" & vbNewLine & _
				"      }" & vbNewLine & _
				"</script>" & vbNewLine

if request("mode") = "" then
	Response.Write	"<form action=""pop_preview.asp"" method=""post"" name=""previewTopic"">" & vbNewLine & _
					"<input type=""hidden"" name=""subject"" value="""">" & vbNewLine & _
					"<input type=""hidden"" name=""message"" value="""">" & vbNewLine & _
					"<input type=""hidden"" name=""sig"" value="""">" & vbNewLine & _
					"<input type=""hidden"" name=""author"" value="""">" & vbNewLine & _
					"<input type=""hidden"" name=""mode"" value=""display"">" & vbNewLine & _
					"</form>" & vbNewLine & _
					"<script language=""JavaScript"" type=""text/javascript"">submitPreview();</script>" & vbNewLine
else
	CColor = strForumCellColor
	strSubjectPreview = trim(Request.Form("subject"))
	strMessagePreview = trim(Request.Form("message"))
	if strMessagePreview = "" or IsNull(strMessagePreview) then
		if strAllowForumCode = "1" then
			strMessagePreview = "[center][b]< There is no text to preview ! >[/b][/center]"
			strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
		else
			strMessagePreview = "<center><b>< There is no text to preview ! ></b></center>"
			strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
		end if
	else
		if Request.Form("author") = "" or isNull(Request.Form("author")) then
			strSigAuthor = strDBNTUserName
		else
			strSigAuthor = ChkString(getMemberName(Request.Form("author")),"SQLString")
		end if
		if Request.Form("sig") = "yes" and trim(GetSig(strSigAuthor)) <> "" then
			if strDSignatures = "1" then
				strMessagePreview = formatStr(chkString(strMessagePreview,"preview")) & "<hr>" & formatStr(chkString(cleancode(GetSig(strSigAuthor)),"preview"))
			else
				strMessagePreview = strMessagePreview & vbNewline & vbNewline & CleanCode(GetSig(strSigAuthor))
				strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
			end if
		else
			strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
		end if
	end if
	if strSubjectPreview = "" or IsNull(strSubjectPreview) then
		strPreviewTitle = "Message Preview"
	else
		strPreviewTitle = "Message Preview - " & chkString(strSubjectPreview,"display")
	end if

	Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td>" & strPreviewTitle & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & strMessagePreview & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine
end if
WriteFooterShort
Response.End
%>
