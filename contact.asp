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
'##   Contact Page MOD v1.1
'#################################################################################
%><!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE file="inc_func_member.asp" --><%
Response.Write "<script type=""text/javascript"" src=""inc_formfieldlimiter.js""></script>" & vbNewLine

If strEmail = "0" Then
	Call FailMessage("<li>The Administrator has turned off all e-mail features.</li>",False)
	WriteFooter
	Response.End
End If

strSql ="SELECT M_NAME, M_USERNAME, M_EMAIL "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
strSql = strSql & " WHERE MEMBER_ID = " & intAdminMemberID & ""
set rs = my_conn.Execute (strSql)

If (rs.EOF or rs.BOF) Then
	Set rs = Nothing
	Call FailMessage("<li>The Administrator's account could not be located</li>",False)
	WriteFooter
	Response.End
Else
	Name = Trim("" & rs("M_NAME"))
	Email = Trim("" & rs("M_EMAIL"))
End If

rs.close
set rs = nothing

If Request.QueryString("mode") = "DoIt" Then
	Err_Msg = ""
	RandCode = Request.Form("code")
	strRCCode = Request.Form("Coder")
	RandCode2 = (strRCCode + 17456) / 50000
	lenCode = Len(RandCode2)
	NullStop = False
	If LenCode < 6 and Nullstop = False then
		For J = 1 to (6 - LenCode)
			NullRC = NullRC & "0"
		Next
		NullStop = True
	End If
	RandCode2 = NullRC & RandCode2
	
	if (Request.Form("YName") = "") then Err_Msg = Err_Msg & "<li>You must enter your name</li>"
	if (Request.Form("YEmail") = "") then
		Err_Msg = Err_Msg & "<li>You must give your email address</li>"
	else
		if (EmailField(Request.Form("YEmail")) = 0) then Err_Msg = Err_Msg & "<li>You Must enter a valid email address</li>"
	end if
	If RandCode <> RandCode2 then Err_Msg = Err_Msg & "<li>Invalid or missing authentication code</li>"
	if (Request.Form("Msg") = "") then Err_Msg = Err_Msg & "<li>You Must enter a message</li>"
	if (Err_Msg = "") then
		strRecipientsName = Name
		strRecipients = Email
		strSubject = strForumTitle
		strMessage = Request.Form("Msg") & vbNewline & vbNewline
		strMessage = strMessage & "You received this from : " & Request.Form("YName") & " (" & Request.Form("YEmail") & ") "
		strMessage = chkBadWords(strMessage)
		strFromName = Request.Form("YName")
		strSender = Request.Form("YEmail")
		
		'Spam filter - Define keywords to filter, then Call the subroutine
		'Const KeyWords = "porn,Viagra,bondage,hardcore,tits,cialis,pussy,penis"
		'Call SpamCheck(strMessage,KeyWords,SPAM)
					
		%><!--#INCLUDE FILE="inc_mail.asp" --><%
		
		'####Start#### Sends a copy of the mail sent from the forum to the sender
		If Request.Form("emailcopy") = "on" Then
			strRecipients = Request.Form("YEmail")
			strFrom = Request.Form("YEmail")
			strSubject = "COPY of Message Sent From " & strForumTitle & " by " & Request.Form("YName")
			strMessage = "Hello " & Request.Form("YName") & vbNewline & vbNewline
			strMessage = strMessage & "Below is a copy of the message you sent to " & strForumTitle & ":"
			strMessage = strMessage & strRName & " " & vbNewline & vbNewline
			strMessage = strMessage & Request.Form("Msg") & vbNewline & vbNewline
			%><!--#INCLUDE FILE="inc_mail.asp" --><%
		end if
		'####End#### Sends a copy of the mail sent from the forum to the sender
		
		Call OkMessage("The administrator has been contacted.","default.asp","Return to the site")
	else
		Call FailMessage(Err_Msg,True)
	end if
else
	Response.Write	"<form action=""contact.asp?mode=DoIt"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input type=""hidden"" name=""Page"" value=""" & Request.QueryString & """>" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Contact the &ldquo;" & strForumTitle & "&rdquo; Administrator</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">All Fields Are Required</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Your Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""YName"" size=""25""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Your E-mail:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""YEmail"" size=""25""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Enter Code:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
		
		strRCCode = Request.QueryString("rc")
		strRC = Request.QueryString("code")
		strRCP = Request.QueryString("p")
		
		If strRC = "image" then
			NullStop = False
			RandCode = (strRCCode + 17456) / 50000
			lenCode = Len(RandCode)
			If LenCode < 6 and Nullstop = False then
				For J = 1 to (6 - LenCode)
					NullRC = NullRC & "0"
				Next
				NullStop = True
			End If
			RandCode = NullRC & RandCode
			ImageP = Mid(RandCode, strRCP,1)
			Response.Redirect "images/" & ImageP & ".gif"
		End If
		
		HowManyNbr=6
		NumbersToShow = ""
		Randomize
		
		For I = 1 to HowManyNbr
			NumbersToShow = NumbersToShow & Fix(9*Rnd)
		Next
		
		RandomizedCode = NumbersToShow * 50000 - 17456
		NullStop = False
		
		For I = 1 to HowManyNbr
			Response.Write  "<img src='contact.asp?code=image&rc=" & RandomizedCode &"&p=" & I & "' border='0' alt='Code'>"
		Next
		
		Response.Write	"&nbsp;&ndash;&nbsp;<input type=""hidden"" name=""Coder"" value=""" & RandomizedCode & """>" & vbNewLine & _
						"<input type=""text"" name=""code"" size=""" & HowManyNbr & """ maxlength=""" & HowManyNbr & """></td>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">Message:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue""><textarea name=""Msg"" id=""msg"" cols=""50"" rows=""10""></textarea><div id=""msg-status""></div></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""formlabel"">&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue""><input type=""Checkbox"" name=""emailcopy""> Send a copy to my email address.</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""options"" colspan=""2""><button type=""submit"" id=""Submit1"" name=""Submit1"">Send</button></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"</table>" & vbNewLine & _
						"</form>" & vbNewLine & _
						"<script type=""text/javascript"">" & vbNewLine & _
						"  fieldlimiter.setup({" & vbNewLine & _
						"  thefield: document.getElementById(""msg"")," & vbNewLine & _
						"  maxlength: 500," & vbNewLine & _
						"  statusids: [""msg-status""]," & vbNewLine & _
						"  onkeypress:function(maxlength, curlength){" & vbNewLine & _
						"}" & vbNewLine & _
						"})" & vbNewLine & _
						"</script>" & vbNewLine
end if

WriteFooter
Response.End

'Spam filter subroutine
Sub SpamCheck(Data,Words,SPAM)
	Dim WordArray, i
	WordArray = Split (Words,",",-1,1)
	For i = 0 to UBound(WordArray)
		If InStr(LCase(Data),LCase(WordArray(i))) Then
			SPAM = True
			Exit For
		End If
	Next
	
	If Trim(Data) = "" or SPAM Then Response.Redirect "http://127.0.0.1/"
End Sub
%>
