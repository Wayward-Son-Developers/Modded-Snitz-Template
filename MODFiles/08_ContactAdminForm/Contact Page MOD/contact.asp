<%
'###############################################################################
'##
'##                   Snitz Forums 2000 v3.4.06
'##
'###############################################################################
'##
'## Copyright © 2000-06 Michael Anderson, Pierre Gorissen,
'##                   Huw Reddick and Richard Kinser
'##
'## This program is free. You can redistribute and/or modify it under the
'## terms of the GNU General Public License as published by the Free Software
'## Foundation; either version 2 or (at your option) any later version.
'##
'## All copyright notices regarding Snitz Forums 2000 must remain intact in
'## the scripts and in the HTML output.  The "powered by" text/logo with a
'## link back to http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful but
'## WITHOUT ANY WARRANTY; without even an implied warranty of MERCHANTABILITY
'## or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License
'## for more details.
'##
'## You should have received a copy of the GNU General Public License along
'## with this program; if not, write to:
'##
'##              Free Software Foundation, Inc.
'##              59 Temple Place, Suite 330
'##              Boston, MA 02111-1307
'##
'## Support can be obtained from our support forums at:
'##
'##              http://forum.snitz.com
'##
'## Correspondence and marketing questions can be sent to:
'##
'##              manderson@snitz.com
'##
'## ***********************************************
'## * Form field Limiter v2.0- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
'## * This notice MUST stay intact for legal use
'## * Visit Project Page at http://www.dynamicdrive.com for full source code
'## ***********************************************
'##
'##   Contact Page MOD v1.1
'##
'###############################################################################
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE file="inc_func_member.asp" -->
<%
   '  ## Forum_SQL
   strSql ="SELECT M_NAME, M_USERNAME, M_EMAIL "
   strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
   strSql = strSql & " WHERE MEMBER_ID = " & intAdminMemberID & ""
   set rs = my_conn.Execute (strSql)
   if (rs.EOF or rs.BOF)then
      Err_Msg = Err_Msg & "<li>The Administrator's account could not be located</li>"
      Response.Write("<p><b><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><center>There Was A Problem</center></font></b></p>" & vbNewLine)
      Response.Write "<table align=""center"">" & vbNewLine & _
         "  <tr>" & vbNewLine & _
         "     <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
         "  </tr>" & vbNewLine & _
         "</table>" & vbNewLine
      set rs = nothing
      Response.Write("<p><font size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick= window.close()"">Close Window</A></font></p>" & vbNewLine)
      Response.End
   else
      Name = Trim("" & rs("M_NAME"))
      Email = Trim("" & rs("M_EMAIL"))
   end if
   rs.close
   set rs = nothing

if Request.QueryString("mode") = "DoIt" then
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

   if (Request.Form("YName") = "") then
      Err_Msg = Err_Msg & "<li>You must enter your name</li>"
   end if
   if (Request.Form("YEmail") = "") then
      Err_Msg = Err_Msg & "<li>You must give your email address</li>"
   else
      if (EmailField(Request.Form("YEmail")) = 0) then
         Err_Msg = Err_Msg & "<li>You Must enter a valid email address</li>"
      end if
   end if
   If RandCode <> RandCode2 then
         Err_Msg = Err_Msg & "<li>Invalid or missing authentication code</li>"
      End If

   if (Request.Form("Msg") = "") then
      Err_Msg = Err_Msg & "<li>You Must enter a message</li>"
   end if
   if lcase(strEmail) = "1" then
      if (Err_Msg = "") then
         strRecipientsName = Name
         strRecipients = Email
         strSubject = strForumTitle
         strMessage = Request.Form("Msg") & vbNewline & vbNewline
         strMessage = strMessage & "You received this from : " & Request.Form("YName") & " (" & Request.Form("YEmail") & ") "
         strFromName = Request.Form("YName")
         strSender = Request.Form("YEmail")

'Spam filter - Define keywords to filter, then Call the subroutine
Const KeyWords = "porn,Viagra,bondage,hardcore,tits,cialis,pussy,penis"
SpamCheck strMessage,KeyWords,SPAM

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

If Trim(Data) = "" or SPAM Then
Response.Redirect "http://pbskids.org/barney/"
End If
End Sub

         %>
         <!--#INCLUDE FILE="inc_mail.asp" -->
         <%

'####Start#### Sends a copy of the mail sent from the forum to the sender
If Request.Form("emailcopy") = "on" Then
strRecipients = Request.Form("YEmail")
strFrom = Request.Form("YEmail")
strSubject = "COPY of Message Sent From " & strForumTitle & " by " & Request.Form("YName")
strMessage = "Hello " & Request.Form("YName") & vbNewline & vbNewline
strMessage = strMessage & "Below is a copy of the message you sent to " & strForumTitle & ":"
strMessage = strMessage & strRName & " " & vbNewline & vbNewline
strMessage = strMessage & Request.Form("Msg") & vbNewline & vbNewline
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
end if
'####End#### Sends a copy of the mail sent from the forum to the sender

         Response.Write("<p><b><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><center><br>" & strForumTitle & " has been contacted</center></font></b></p>" & vbNewLine)
         Response.write " <p align=""center""><font size=""" & strDefaultFontSize & """ align=""center""><a href=""default.asp"">Visit the Forum</a></font></p><br /><br />" & vbNewLine
      else
         Response.Write("<p><b><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><center><br>There Was A Problem</center></font></b></p>" & vbNewLine)
         Response.Write "<table align=""center"">" & vbNewLine & _
            "<tr>" & vbNewLine & _
            "<td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
            "</tr>" & vbNewLine & _
            "</table>" & vbNewLine
         Response.Write("<p><font size=""" & strDefaultFontSize & """><center><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></center></font></p><br><br>" & vbNewLine)
      end if
   end if
else
   Response.Write "<form action=""contact.asp?mode=DoIt"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
      "  <input type=""hidden"" name=""Page"" value=""" & Request.QueryString & """>" & vbNewLine & _
      "  <br><table align=""center"" border=""0"" width=""95%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
      "     <tr>" & vbNewLine & _
      "        <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
      "           <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
      "              <tr>" & vbNewLine & _
      "                 <td colspan=""2"" bgColor=""" & strHeadCellColor & """ align=""center"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Contact " & strForumTitle & "</font></b></td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _

      "              <tr>" & vbNewLine & _
      "                 <td colspan=""2"" bgColor=""" & strCategoryCellColor & """ align=""center"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>All Fields Are Required</font></td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _


      "              <tr>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Name:</font></b></td>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""YName"" size=""25""></td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _
      "              <tr>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your E-mail:</font></b></td>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""YEmail"" size=""25""></td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _


      "              <tr>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Enter Code:</font></b></td>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine
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
         Response.Write  "    <img src='contact.asp?code=image&rc=" & RandomizedCode &"&p=" & I & "' border='0' alt='Code'>"
      Next
   Response.Write "   <input type=""hidden"" name=""Coder"" value=""" & RandomizedCode & """>" & vbNewLine & _
                  "   <input type=""text"" name=""code"" size=""" & HowManyNbr & """ maxlength=""" & HowManyNbr & """></td>" & vbNewLine & _
      "              <tr>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """ align=""right"" valign=""top"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Message:</font></b></td>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """><textarea name=""Msg"" id=""msg"" cols=""50"" rows=""10""></textarea><div><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ id=""msg-status""></div></td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _

      "              <tr>" & vbNewLine & _
      "              <td bgColor=""" & strPopUpTableColor & """></td>" & vbNewLine & _
      "              <td bgColor=""" & strPopUpTableColor & """ align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><input type=""Checkbox"" name=""emailcopy""> Send a copy to my email address" & vbnewline & _
      "              </td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _

      "              <tr>" & vbNewLine & _
      "                 <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
      "              </tr>" & vbNewLine & _
      "           </table>" & vbNewLine & _
      "        </td>" & vbNewLine & _
      "     </tr>" & vbNewLine & _
      "  </table>" & vbNewLine & _
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
%>
