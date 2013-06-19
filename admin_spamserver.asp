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
If Session(strCookieURL & "Approval") <> strAdminCode Then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
End If

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Blocked E-Mail Domains</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

Dim strMethodType
If Request.Form("Method_Type") <> "" Then
	strMethodType = LCase(Trim(Request.Form("Method_Type")))
Else
	If Request.QueryString("Method_Type") <> "" Then
		strMethodType = LCase(Trim(Request.QueryString("Method_Type")))
	Else
		strMethodType = "blank"
	End If
End If

Select Case strMethodType
	Case "add"
		Dim strSpammServer : strSpammServer = LCase(Trim(chkString(Request.Form("SpamServer"),"sqlstring")))
		
		Err_Msg = ""
		
		if strSpammServer = "" then
			Err_Msg = Err_Msg & "<li>You need to enter an address to block.</li>"
		end if
		
		if (Instr(strSpammServer, " ") > 0 ) then
			Err_Msg = Err_Msg & "<li>You cannot have spaces in your address.</li>"
		end if
		
		'Comment out down to the next comment to let it take me@example.com and/or .ex as well
		If Left(strSpammServer,1) = "@" Then
			If InStr(1,strSpammServer,".",vbTextCompare) = 0 Then
				Err_Msg = Err_Msg & "<li>You need to have a TLD (.com, .net, .whatever) in your address.</li>"
			End If
		Else
			If InStr(1,strSpammServer,"@",vbTextCompare) <> 0 Then
				Err_Msg = Err_Msg & "<li>You can only enter a domain (@example.com), not a specific address (me@example.com).</li>"
			Else
				If InStr(1,strSpammServer,".",vbTextCompare) = 0 Then
					Err_Msg = Err_Msg & "<li>You need to enter a valid domain (@example.com).</li>"
				Else
					strSpammServer = "@" & strSpammServer
				End If
			End If
		End if
		'Comment out up to the previous comment to let it take me@example.com and/or .ex as well
		
		If Err_Msg = "" Then
			strSQL = "SELECT SPAM_ID, SPAM_SERVER FROM " & strFilterTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '"& strSpammServer &"'"
			set rs = my_conn.execute(strSQL)
			
			If Not rs.EOF And Not rs.BOF Then
				Err_Msg = Err_Msg & "<li>'" & strSpammServer & "' is already in the list.</li>"
			End If
			
			rs.Close
			Set rs = Nothing
		End If
		
		if Err_Msg = "" then
			strSQL = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('" & strSpammServer & "')"
			my_Conn.Execute (strSql)
			Call OkMessage("E-Mail domain added","admin_spamserver.asp","Back to the list")
		else
			Call FailMessage(Err_Msg,True)
		end if
	Case "delete"
		If Request.QueryString("id") <> "" And IsNumeric(Request.QueryString("id")) Then
			Dim intSpamID : intSpamID = cLng(Request.QueryString("id"))
			
			strSQL = "DELETE FROM " & strFilterTablePrefix & "SPAM_MAIL WHERE SPAM_ID = " & intSpamID
			my_Conn.Execute (strSql)
			
			Call OkMessage("E-Mail domain deleted","admin_spamserver.asp","Back to the list")
		Else
			Call FailMessage("<li>The domain ID was not numeric.</li>",True)
		End If
	Case Else
		Response.Write	"<table class=""admin"" width=""60%"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td colspan=""2"">Blocked E-Mail Domains</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr valign=""top"">" & vbNewLine & _
						"<td class=""formlabel"">Insert E-Mail Domain:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & _
						"<form action=""admin_spamserver.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
						"<input type=""hidden"" name=""Method_Type"" value=""add"">" & vbNewLine & _
						"@<input  type=""text"" name=""SpamServer"" size=""50"" maxlength=""255"" value="""">&nbsp;" & _
						"<button type=""submit"" id=""submit1"" name=""submit1"">Block</button></td>" & vbNewLine & _
						"</form>"  & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">E-Mail Domains<br />currently blocked:&nbsp;</td>"& vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine
		
		Set rs = my_conn.Execute("SELECT SPAM_ID, SPAM_SERVER FROM " & strFilterTablePrefix & "SPAM_MAIL ORDER BY SPAM_SERVER")
		
		If rs.EOF or rs.BOF Then
			Response.Write	"No blocked e-mail domains found." & vbNewLine
		Else
			Response.Write	"<ul>" & vbNewLine
			Do Until rs.EOF
				Response.Write	"<li>" & rs("SPAM_SERVER") & "&nbsp;" & _
								"<a href=""#"" onClick=""JavaScript:if(window.confirm('Delete Spam Server on this domain?')){window.location=('admin_spamserver.asp?Method_Type=delete&id=" & rs("SPAM_ID") & "');}"">" & _
								getCurrentIcon(strIconTrashcan,"Delete Spam Server address","") & "</a></li>" & vbNewLine
				rs.MoveNext
			Loop
			Response.Write	"</ul>" & vbNewLine
		End If
		
		rs.Close
		Set rs = Nothing
		Response.Write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine
End Select

WriteFooter
%>