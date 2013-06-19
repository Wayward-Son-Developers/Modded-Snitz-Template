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
<!--#INCLUDE FILE="inc_func_member.asp" -->
<% 
if strLogonForMail = "1" and mlev = 0 then
	Call FailMessage("<li>You must be logged on to send a message</li>",False)
	WriteFooterShort
	Response.End
end if

if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
	intMemberID = cLng(Request.QueryString("ID"))
else
	intMemberID = 0
end if

if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""
	
	strSql = "SELECT M_NAME, M_POSTS, M_ALLOWEMAIL FROM " & strMemberTablePrefix & "MEMBERS M"
	strSql = strSql & " WHERE M.MEMBER_ID = " & MemberID
	
	set rs = my_Conn.Execute (strSql)
	
	If Not rs.EOF then
		If Not IsNumeric(rs("M_POSTS")) Then
			intMPosts = 0
		Else
			intMPosts = cLng(rs("M_POSTS"))
		End If
		If Not IsNumeric(rs("M_ALLOWEMAIL")) Then
			intAllowEmail = 0
		Else
			intAllowEmail = cInt(rs("M_ALLOWEMAIL"))
		End If
		
		If intMPosts < intMaxPostsToEMail and intAllowEmail <> "1" Then
			Err_Msg = "<li>" & strNoMaxPostsToEMail & "</li>"
			strSpammerName = RS("M_NAME")
			rs.Close
			
			strSql = "SELECT M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
			strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID
			
			set rs = my_Conn.Execute (strSql)
			
			If rs.bof or rs.eof Then
				strDestName = ""
			Else
				strDestname = rs("M_NAME")
			End If
			rs.close
			
			'Send email to forum admin
			strRecipients = strSender
			strFrom = strSender
			strFromName = "Automatic Server Email"
			strSubject = "Possible Spam Poster"
			strMessage = "There is a possible spam poster at " & strForumTitle & vbNewLine & vbNewLine
			strMessage = strMessage & "Member " & strSpammerName & ", with MemberID " & MemberID & ", has been trying to send emails to " & strDestName & ", without having enough posts to be allowed to do it." & vbNewLine & vbNewLine
			strMessage = strMessage & "He has " & intMPosts & " posts, and should have " & intMaxPostsToEMail & " posts." & vbNewLine & vbNewLine
			strMessage = strMessage & "Here are the message contents: " & VbNewLine & Request.Form("Msg") & vbNewLine & vbNewLine & vbNewLine & vbNewLine
			strMessage = strMessage & "This is a message sent automatically by the Spam Control Mod."
			%><!--#INCLUDE FILE="inc_mail.asp" --><%
		End If
	Else
		rs.Close
	End If
End If

'## Forum_SQL
strSql = "SELECT M.M_RECEIVE_EMAIL, M.M_EMAIL, M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID

set rs = my_Conn.Execute (strSql)

if rs.bof or rs.eof then
	rs.close
	set rs = nothing
	Call FailMessage("<li>There is no Member with that Member ID</li>",False)
else
	strRName = ChkString(rs("M_NAME"),"display")
	strREmail = rs("M_EMAIL")
	strRReceiveEmail = rs("M_RECEIVE_EMAIL")
	
	rs.close
	set rs = nothing
	
	if mLev > 2 or strRReceiveEmail = "1" then
		if lcase(strEmail) = "1" then
			if Request.QueryString("mode") = "DoIt" then
				if mLev => 2 then
					strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
					strSql = strSql & " WHERE MEMBER_ID = " & MemberID

					set rs2 = my_conn.Execute (strSql)
					YName = rs2("M_NAME")
					YEmail = rs2("M_EMAIL")
					set rs2 = nothing
				else
					YName = Request.Form("YName")
					YEmail = Request.Form("YEmail")
					if YName = "" then
						Err_Msg = Err_Msg & "<li>You must enter your UserName</li>"
					end if
					if YEmail = "" then 
						Err_Msg = Err_Msg & "<li>You must give your e-mail address</li>"
					else
						if EmailField(YEmail) = 0 then 
							Err_Msg = Err_Msg & "<li>You must enter a valid e-mail address</li>"
						end if
					end if
				end if
				if Request.Form("Msg") = "" then 
					Err_Msg = Err_Msg & "<li>You must enter a message</li>"
				end if
				'##  E-mails Message to the Author of this Reply.  
				if (Err_Msg = "") then
					strRecipientsName = strRName
					strRecipients = strREmail
					strFrom = YEmail
					strFromName = YName
					strSubject = "Sent From " & strForumTitle & " by " & YName
					strMessage = "Hello " & strRName & vbNewline & vbNewline
					strMessage = strMessage & "You received the following message from: " & YName & " (" & YEmail & ") " & vbNewline & vbNewline 
					strMessage = strMessage & "At: " & strForumURL & vbNewline & vbNewline
					strMessage = strMessage & Request.Form("Msg") & vbNewline & vbNewline

					if strFrom <> "" then 
						strSender = strFrom
					end if
					%><!--#INCLUDE FILE="inc_mail.asp" --><%
					Call PopOkMessage("E-mail has been sent",False)
				else
					Call FailMessage(Err_Msg,True)
					WriteFooterShort
					Response.End 
				end if
			else 
				Err_Msg = ""
				if trim(strREmail) <> "" then
					strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
					strSql = strSql & " WHERE MEMBER_ID = " & MemberID

					set rs2 = my_conn.Execute (strSql)
					YName = ""
					YEmail = ""

					if (rs2.EOF or rs2.BOF)  then
						if strLogonForMail <> "0" then 
							Err_Msg = Err_Msg & "<li>You must be logged on to send a message</li>"
							Call FailMessage(Err_Msg,False)
							WriteFooterShort
							Response.End
						end if
					else
						YName = Trim("" & rs2("M_NAME"))
						YEmail = Trim("" & rs2("M_EMAIL"))
					end if
					rs2.close
					set rs2 = nothing

					Response.Write	"<form action=""pop_mail.asp?mode=DoIt&id=" & intMemberID & """ method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
									"<table class=""content"" width=""100%"">" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td class=""formlabel"">Send To:&nbsp;</td>" & vbNewLine & _
									"<td class=""formvalue"">" & strRName & "</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td class=""formlabel"">Your Name:&nbsp;</td>" & vbNewLine & _
									"<td class=""formvalue"">"
					if YName = "" then
						Response.Write "<input name=""YName"" type=""text"" value=""" & YName & """ size=""25"">"
					else
						Response.Write YName
					end if
					Response.Write	"</td></tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td class=""formlabel"">Your E-mail:&nbsp;</td>" & vbNewLine & _
									"<td class=""formvalue"">"
					if YEmail = "" then
						Response.Write "<input name=""YEmail"" type=""text"" value=""" & YEmail & """ size=""25"">"
					else
						Response.Write YEmail
					end if
					Response.Write	"</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td class=""formlabel"" colspan=""2"">Message:&nbsp;</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td colspan=""2""><textarea name=""Msg"" cols=""40"" rows=""5""></textarea></td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td class=""options"" colspan=""2""><button type=""Submit"" id=""Submit1"" name=""Submit1"">Send</button></td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"</table>" & vbNewLine & _
									"</form>" & vbNewLine
				else
					Call FailMessage("<li>No E-mail address is available for this user.</li>",False)
				end if
			end if
		else
			Response.Write	"<p class=""content"">Click to send <a href=""mailto:" & chkString(strREmail,"display") & """>" & strRName & "</a> an e-mail</p>" & vbNewLine
		end if
	else
		Response.Write	"<p class=""content"">This Member does not wish to receive e-mail.</p>" & vbNewLine
	end if
end if
WriteFooterShort
Response.End
%>
