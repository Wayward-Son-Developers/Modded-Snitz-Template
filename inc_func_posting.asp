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


function GetKey(action)
	'// Create an array of characters to choose from for the key.
	'// If you would like to add uppercase letters or high ASCII characters,
	'// simply add them to the array, just remember to modify intNumChars
	'// variable to match number of characters in the array.

	intNumChars = 52

	keyArray = Array("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","0","1","2","3","4","5","6","7","8","9","0","1","2","3","4","5","6","7","8","9","0","1","2","3","4","5")

	'// This picks 32 random numbers and pulls corresponding letters from the
	'// array. If you want a larger, or smaller key, simply adjust the
	'// number of characters you grab.

	dim key(32)
	Randomize
	for i = 0 to 31
		key(i) = (Int(((intNumChars - 1) * Rnd) + 1))
	next

	'// Make the key!

	strKey = ""
	for j = 0 to 31
		strKey = strKey & keyArray(key(j))
	next

	GetKey = strKey
	if action = "sendemail" then
		'## E-mails verification link to the new e-mail address.
		strRecipientsName = Request.Form("Name")
		strRecipients = Request.Form("Email")
		strFrom = strSender
		strFromName = strForumTitle
		strsubject = strForumTitle & "- Your E-mail Address Has Been Changed "
		strMessage = "Hello " & Request.Form("name") & vbNewLine & vbNewLine
		if Request.QueryString("mode") <> "EditIt" then
			strMessage = strMessage & "You received this message from " & strForumTitle & " because someone has changed your e-mail address on the forums at " & strForumURL & vbNewLine & vbNewLine
		else
			strMessage = strMessage & "You received this message from " & strForumTitle & " because you have changed your e-mail address on the forums at " & strForumURL & vbNewLine & vbNewLine
		end if
		strMessage = strMessage & "To complete your e-mail change, please click on the link below:" & vbNewLine & vbNewLine
		strMessage = strMessage & strForumURL & "pop_profile.asp?verkey=" & strKey & vbNewLine & vbNewLine
		strMessage = strMessage & "Thank You!" & vbNewLine & vbNewLine
		strMessage = strMessage & "Forum Admin"
%>
		<!--#INCLUDE FILE="inc_mail.asp" -->
<%
	end if
end function

function CleanCode(fString)
	'## left for compatibility with older versions of the forum
	if fString = "" or IsNull(fstring) then 
		fString = " "
	else
		fString = replace(fString, "<BLOCKQUOTE id=quote><font size=" & strFooterFontSize & " face=""" & strDefaultFontFace & """ id=quote>quote:<hr height=1 noshade id=quote>","[quote]", 1, -1, 1)
		fString = replace(fString, "<hr height=1 noshade id=quote></BLOCKQUOTE id=quote></font id=quote><font face=""" & strDefaultFontFace & """ size=" & strDefaultFontSize & " id=quote>","[/quote]", 1, -1, 1)
		fString = replace(fString, "<blockquote id=""quote""><font size=""" & strFooterFontSize & """ face=""" & strDefaultFontFace & """ id=""quote"">quote:<hr height=""1"" noshade id=""quote"">","[quote]", 1, -1, 1)
		fString = replace(fString, "<hr height=""1"" noshade id=""quote""></font id=""quote""></blockquote id=""quote"">","[/quote]", 1, -1, 1)
		fString = replace(fString, "<hr height=""1"" noshade id=""quote""></blockquote id=""quote""></font id=""quote"">","[/quote]", 1, -1, 1)
		if strAllowForumCode = "1" then
			fString = replace(fString, "<b>","[b]", 1, -1, 1)
			fString = replace(fString, "</b>","[/b]", 1, -1, 1)
			fString = replace(fString, "<s>", "[s]", 1, -1, 1)
		    fString = replace(fString, "</s>", "[/s]", 1, -1, 1)
			fString = replace(fString, "<u>","[u]", 1, -1, 1)
			fString = replace(fString, "</u>","[/u]", 1, -1, 1)
			fString = replace(fString, "<i>","[i]", 1, -1, 1)
			fString = replace(fString, "</i>","[/i]", 1, -1, 1)
			fString = replace(fString, "<font face='Andale Mono'>", "[font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "</font id='Andale Mono'>", "[/font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "<font face='Arial'>", "[font=Arial]", 1, -1, 1)
			fString = replace(fString, "</font id='Arial'>", "[/font=Arial]", 1, -1, 1)
			fString = replace(fString, "<font face='Arial Black'>", "[font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "</font id='Arial Black'>", "[/font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "<font face='Book Antiqua'>", "[font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "</font id='Book Antiqua'>", "[/font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "<font face='Century Gothic'>", "[font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "</font id='Century Gothic'>", "[/font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "<font face='Comic Sans MS'>", "[font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "</font id='Comic Sans MS'>", "[/font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "<font face='Courier New'>", "[font=Courier New]", 1, -1, 1)
			fString = replace(fString, "</font id='Courier New'>", "[/font=Courier New]", 1, -1, 1)
			fString = replace(fString, "<font face='Georgia'>", "[font=Georgia]", 1, -1, 1)
			fString = replace(fString, "</font id='Georgia'>", "[/font=Georgia]", 1, -1, 1)
			fString = replace(fString, "<font face='Impact'>", "[font=Impact]", 1, -1, 1)
			fString = replace(fString, "</font id='Impact'>", "[/font=Impact]", 1, -1, 1)
			fString = replace(fString, "<font face='Tahoma'>", "[font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "</font id='Tahoma'>", "[/font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "<font face='Times New Roman'>", "[font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "</font id='Times New Roman'>", "[/font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "<font face='Trebuchet MS'>", "[font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "</font id='Trebuchet MS'>", "[/font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "<font face='Script MT Bold'>", "[font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "</font id='Script MT Bold'>", "[/font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "<font face='Stencil'>", "[font=Stencil]", 1, -1, 1)
			fString = replace(fString, "</font id='Stencil'>", "[/font=Stencil]", 1, -1, 1)
			fString = replace(fString, "<font face='Verdana'>", "[font=Verdana]", 1, -1, 1)
			fString = replace(fString, "</font id='Verdana'>", "[/font=Verdana]", 1, -1, 1)
			fString = replace(fString, "<font face='Lucida Console'>", "[font=Lucida Console]", 1, -1, 1)
			fString = replace(fString, "</font id='Lucida Console'>", "[/font=Lucida Console]", 1, -1, 1)
			fString = replace(fString, "<font face=""Andale Mono"">", "[font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "</font id=""Andale Mono"">", "[/font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "<font face=""Arial"">", "[font=Arial]", 1, -1, 1)
			fString = replace(fString, "</font id=""Arial"">", "[/font=Arial]", 1, -1, 1)
			fString = replace(fString, "<font face=""Arial Black"">", "[font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "</font id=""Arial Black"">", "[/font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "<font face=""Book Antiqua"">", "[font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "</font id=""Book Antiqua"">", "[/font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "<font face=""Century Gothic"">", "[font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "</font id=""Century Gothic"">", "[/font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "<font face=""Comic Sans MS"">", "[font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "</font id=""Comic Sans MS"">", "[/font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "<font face=""Courier New"">", "[font=Courier New]", 1, -1, 1)
			fString = replace(fString, "</font id=""Courier New"">", "[/font=Courier New]", 1, -1, 1)
			fString = replace(fString, "<font face=""Georgia"">", "[font=Georgia]", 1, -1, 1)
			fString = replace(fString, "</font id=""Georgia"">", "[/font=Georgia]", 1, -1, 1)
			fString = replace(fString, "<font face=""Impact"">", "[font=Impact]", 1, -1, 1)
			fString = replace(fString, "</font id=""Impact"">", "[/font=Impact]", 1, -1, 1)
			fString = replace(fString, "<font face=""Tahoma"">", "[font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "</font id=""Tahoma"">", "[/font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "<font face=""Times New Roman"">", "[font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "</font id=""Times New Roman"">", "[/font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "<font face=""Trebuchet MS"">", "[font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "</font id=""Trebuchet MS"">", "[/font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "<font face=""Script MT Bold"">", "[font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "</font id=""Script MT Bold"">", "[/font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "<font face=""Stencil"">", "[font=Stencil]", 1, -1, 1)
			fString = replace(fString, "</font id=""Stencil"">", "[/font=Stencil]", 1, -1, 1)
			fString = replace(fString, "<font face=""Verdana"">", "[font=Verdana]", 1, -1, 1)
			fString = replace(fString, "</font id=""Verdana"">", "[/font=Verdana]", 1, -1, 1)
			fString = replace(fString, "<font face=""Lucida Console"">", "[font=Lucida Console]", 1, -1, 1)
			fString = replace(fString, "</font id=""Lucida Console"">", "[/font=Lucida Console]", 1, -1, 1)
			fString = replace(fString, "<font color=red>", "[red]", 1, -1, 1)
			fString = replace(fString, "</font id=red>", "[/red]", 1, -1, 1)
			fString = replace(fString, "<font color=green>", "[green]", 1, -1, 1)
			fString = replace(fString, "</font id=green>", "[/green]", 1, -1, 1)
			fString = replace(fString, "<font color=blue>", "[blue]", 1, -1, 1)
			fString = replace(fString, "</font id=blue>", "[/blue]", 1, -1, 1)
			fString = replace(fString, "<font color=white>", "[white]", 1, -1, 1)
			fString = replace(fString, "</font id=white>", "[/white]", 1, -1, 1)
			fString = replace(fString, "<font color=purple>", "[purple]", 1, -1, 1)
			fString = replace(fString, "</font id=purple>", "[/purple]", 1, -1, 1)
			fString = replace(fString, "<font color=yellow>", "[yellow]", 1, -1, 1)
			fString = replace(fString, "</font id=yellow>", "[/yellow]", 1, -1, 1)
			fString = replace(fString, "<font color=violet>", "[violet]", 1, -1, 1)
			fString = replace(fString, "</font id=violet>", "[/violet]", 1, -1, 1)
			fString = replace(fString, "<font color=brown>", "[brown]", 1, -1, 1)
			fString = replace(fString, "</font id=brown>", "[/brown]", 1, -1, 1)
			fString = replace(fString, "<font color=black>", "[black]", 1, -1, 1)
			fString = replace(fString, "</font id=black>", "[/black]", 1, -1, 1)
			fString = replace(fString, "<font color=pink>", "[pink]", 1, -1, 1)
			fString = replace(fString, "</font id=pink>", "[/pink]", 1, -1, 1)
			fString = replace(fString, "<font color=orange>", "[orange]", 1, -1, 1)
			fString = replace(fString, "</font id=orange>", "[/orange]", 1, -1, 1)
			fString = replace(fString, "<font color=gold>", "[gold]", 1, -1, 1)
			fString = replace(fString, "</font id=gold>", "[/gold]", 1, -1, 1)
			fString = replace(fString, "<font color=beige>", "[beige]", 1, -1, 1)
			fString = replace(fString, "</font id=beige>", "[/beige]", 1, -1, 1)
			fString = replace(fString, "<font color=teal>", "[teal]", 1, -1, 1)
			fString = replace(fString, "</font id=teal>", "[/teal]", 1, -1, 1)
			fString = replace(fString, "<font color=navy>", "[navy]", 1, -1, 1)
			fString = replace(fString, "</font id=navy>", "[/navy]", 1, -1, 1)
			fString = replace(fString, "<font color=maroon>", "[maroon]", 1, -1, 1)
			fString = replace(fString, "</font id=maroon>", "[/maroon]", 1, -1, 1)
			fString = replace(fString, "<font color=limegreen>", "[limegreen]", 1, -1, 1)
			fString = replace(fString, "</font id=limegreen>", "[/limegreen]", 1, -1, 1)
			fString = replace(fString, "<font color=""red"">", "[red]", 1, -1, 1)
			fString = replace(fString, "</font id=""red"">", "[/red]", 1, -1, 1)
			fString = replace(fString, "<font color=""green"">", "[green]", 1, -1, 1)
			fString = replace(fString, "</font id=""green"">", "[/green]", 1, -1, 1)
			fString = replace(fString, "<font color=""blue"">", "[blue]", 1, -1, 1)
			fString = replace(fString, "</font id=""blue"">", "[/blue]", 1, -1, 1)
			fString = replace(fString, "<font color=""white"">", "[white]", 1, -1, 1)
			fString = replace(fString, "</font id=""white"">", "[/white]", 1, -1, 1)
			fString = replace(fString, "<font color=""purple"">", "[purple]", 1, -1, 1)
			fString = replace(fString, "</font id=""purple"">", "[/purple]", 1, -1, 1)
			fString = replace(fString, "<font color=""yellow"">", "[yellow]", 1, -1, 1)
			fString = replace(fString, "</font id=""yellow"">", "[/yellow]", 1, -1, 1)
			fString = replace(fString, "<font color=""violet"">", "[violet]", 1, -1, 1)
			fString = replace(fString, "</font id=""violet"">", "[/violet]", 1, -1, 1)
			fString = replace(fString, "<font color=""brown"">", "[brown]", 1, -1, 1)
			fString = replace(fString, "</font id=""brown"">", "[/brown]", 1, -1, 1)
			fString = replace(fString, "<font color=""black"">", "[black]", 1, -1, 1)
			fString = replace(fString, "</font id=""black"">", "[/black]", 1, -1, 1)
			fString = replace(fString, "<font color=""pink"">", "[pink]", 1, -1, 1)
			fString = replace(fString, "</font id=""pink"">", "[/pink]", 1, -1, 1)
			fString = replace(fString, "<font color=""orange"">", "[orange]", 1, -1, 1)
			fString = replace(fString, "</font id=""orange"">", "[/orange]", 1, -1, 1)
			fString = replace(fString, "<font color=""gold"">", "[gold]", 1, -1, 1)
			fString = replace(fString, "</font id=""gold"">", "[/gold]", 1, -1, 1)
			fString = replace(fString, "<font color=""beige"">", "[beige]", 1, -1, 1)
			fString = replace(fString, "</font id=""beige"">", "[/beige]", 1, -1, 1)
			fString = replace(fString, "<font color=""teal"">", "[teal]", 1, -1, 1)
			fString = replace(fString, "</font id=""teal"">", "[/teal]", 1, -1, 1)
			fString = replace(fString, "<font color=""navy"">", "[navy]", 1, -1, 1)
			fString = replace(fString, "</font id=""navy"">", "[/navy]", 1, -1, 1)
			fString = replace(fString, "<font color=""maroon"">", "[maroon]", 1, -1, 1)
			fString = replace(fString, "</font id=""maroon"">", "[/maroon]", 1, -1, 1)
			fString = replace(fString, "<font color=""limegreen"">", "[limegreen]", 1, -1, 1)
			fString = replace(fString, "</font id=""limegreen"">", "[/limegreen]", 1, -1, 1)
			fString = replace(fString, "<h1>", "[h1]", 1, -1, 1)
			fString = replace(fString, "</h1>", "[/h1]", 1, -1, 1)
			fString = replace(fString, "<h2>", "[h2]", 1, -1, 1)
			fString = replace(fString, "</h2>", "[/h2]", 1, -1, 1)
			fString = replace(fString, "<h3>", "[h3]", 1, -1, 1)
			fString = replace(fString, "</h3>", "[/h3]", 1, -1, 1)
			fString = replace(fString, "<h4>", "[h4]", 1, -1, 1)
			fString = replace(fString, "</h4>", "[/h4]", 1, -1, 1)
			fString = replace(fString, "<h5>", "[h5]", 1, -1, 1)
			fString = replace(fString, "</h5>", "[/h5]", 1, -1, 1)
			fString = replace(fString, "<h6>", "[h6]", 1, -1, 1)
			fString = replace(fString, "</h6>", "[/h6]", 1, -1, 1)
			fString = replace(fString, "<font size=1>", "[size=1]", 1, -1, 1)
			fString = replace(fString, "</font id=size1>", "[/size=1]", 1, -1, 1)
			fString = replace(fString, "<font size=2>", "[size=2]", 1, -1, 1)
			fString = replace(fString, "</font id=size2>", "[/size=2]", 1, -1, 1)
			fString = replace(fString, "<font size=3>", "[size=3]", 1, -1, 1)
			fString = replace(fString, "</font id=size3>", "[/size=3]", 1, -1, 1)
			fString = replace(fString, "<font size=4>", "[size=4]", 1, -1, 1)
			fString = replace(fString, "</font id=size4>", "[/size=4]", 1, -1, 1)
			fString = replace(fString, "<font size=5>", "[size=5]", 1, -1, 1)
			fString = replace(fString, "</font id=size5>", "[/size=5]", 1, -1, 1)
			fString = replace(fString, "<font size=6>", "[size=6]", 1, -1, 1)
			fString = replace(fString, "</font id=size6>", "[/size=6]", 1, -1, 1)
			fString = replace(fString, "<font size=""1"">", "[size=1]", 1, -1, 1)
			fString = replace(fString, "</font id=""size1"">", "[/size=1]", 1, -1, 1)
			fString = replace(fString, "<font size=""2"">", "[size=2]", 1, -1, 1)
			fString = replace(fString, "</font id=""size2"">", "[/size=2]", 1, -1, 1)
			fString = replace(fString, "<font size=""3"">", "[size=3]", 1, -1, 1)
			fString = replace(fString, "</font id=""size3"">", "[/size=3]", 1, -1, 1)
			fString = replace(fString, "<font size=""4"">", "[size=4]", 1, -1, 1)
			fString = replace(fString, "</font id=""size4"">", "[/size=4]", 1, -1, 1)
			fString = replace(fString, "<font size=""5"">", "[size=5]", 1, -1, 1)
			fString = replace(fString, "</font id=""size5"">", "[/size=5]", 1, -1, 1)
			fString = replace(fString, "<font size=""6"">", "[size=6]", 1, -1, 1)
			fString = replace(fString, "</font id=""size6"">", "[/size=6]", 1, -1, 1)
			fString = replace(fString, "<br />","[br]", 1, -1, 1)
			fString = replace(fString, "<hr noshade size=""1"">","[hr]", 1, -1, 1)
			fString = replace(fString, "<div align=left>", "[left]", 1, -1, 1)
		    fString = replace(fString, "</div id=left>", "[/left]", 1, -1, 1)
			fString = replace(fString, "<div align=""left"">", "[left]", 1, -1, 1)
		    fString = replace(fString, "</div id=""left"">", "[/left]", 1, -1, 1)
			fString = replace(fString, "<center>","[center]", 1, -1, 1)
			fString = replace(fString, "</center>","[/center]", 1, -1, 1)
		    fString = replace(fString, "<div align=right>", "[right]", 1, -1, 1)
		    fString = replace(fString, "</div id=right>", "[/right]", 1, -1, 1)
			fString = replace(fString, "<div align=""right"">", "[right]", 1, -1, 1)
			fString = replace(fString, "</div id=""right"">", "[/right]", 1, -1, 1)
			fString = replace(fString, "<ul>","[list]", 1, -1, 1)
			fString = replace(fString, "</ul>","[/list]", 1, -1, 1)
			fString = replace(fString, "<ol type=1>","[list=1]", 1, -1, 1)
			fString = replace(fString, "</ol id=1>","[/list=1]", 1, -1, 1)
			fString = replace(fString, "<ol type=a>","[list=a]", 1, -1, 1)
			fString = replace(fString, "</ol id=a>","[/list=a]", 1, -1, 1)
			fString = replace(fString, "<ol type=""1"">","[list=1]", 1, -1, 1)
			fString = replace(fString, "</ol id=""1"">","[/list=1]", 1, -1, 1)
			fString = replace(fString, "<ol type=""a"">","[list=a]", 1, -1, 1)
			fString = replace(fString, "</ol id=""a"">","[/list=a]", 1, -1, 1)
			fString = replace(fString, "<li>","[*]", 1, -1, 1)
			fString = replace(fString, "</li>","[/*]", 1, -1, 1)
			fString = replace(fString, "<pre id=code><font face=courier size=" & strDefaultFontSize & " id=code>","[code]", 1, -1, 1)
			fString = replace(fString, "</font id=code></pre id=code>","[/code]", 1, -1, 1)
			fString = replace(fString, "<pre id=""code""><font face=""courier"" size=""" & strDefaultFontSize & """ id=""code"">","[code]", 1, -1, 1)
			fString = replace(fString, "</font id=""code""></pre id=""code"">","[/code]", 1, -1, 1)
			fString = replace(fString, "<div id=""scrollcode"" class=""scrollcode""><pre>","[scrollcode]", 1, -1, 1)
			fString = replace(fString, "</pre></div id=""scrollcode"">","[/scrollcode]", 1, -1, 1)
		end if
		if strIcons = "1" then
			fString = replace(fString, "<img src=icon_smile_angry.gif border=0 align=middle>", "[:(!]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_blackeye.gif border=0 align=middle>", "[B)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_dead.gif border=0 align=middle>", "[xx(]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_dead.gif border=0 align=middle>", "[XX(]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_shock.gif border=0 align=middle>", "[:O]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_shock.gif border=0 align=middle>", "[:o]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_shock.gif border=0 align=middle>", "[:0]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_blush.gif border=0 align=middle>", "[:I]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_sad.gif border=0 align=middle>", "[:(]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_shy.gif border=0 align=middle>", "[8)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile.gif border=0 align=middle>", "[:)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_evil.gif border=0 align=middle>", "[}:)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_big.gif border=0 align=middle>", "[:D]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_cool.gif border=0 align=middle>", "[8D]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_sleepy.gif border=0 align=middle>", "[|)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_clown.gif border=0 align=middle>", "[:o)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_clown.gif border=0 align=middle>", "[:O)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_clown.gif border=0 align=middle>", "[:0)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_tongue.gif border=0 align=middle>", "[:P]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_tongue.gif border=0 align=middle>", "[:p]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_wink.gif border=0 align=middle>", "[;)]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_8ball.gif border=0 align=middle>", "[8]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_question.gif border=0 align=middle>", "[?]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_approve.gif border=0 align=middle>", "[^]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_disapprove.gif border=0 align=middle>", "[V]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_disapprove.gif border=0 align=middle>", "[v]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_dissapprove.gif border=0 align=middle>", "[V]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_dissapprove.gif border=0 align=middle>", "[v]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_kisses.gif border=0 align=middle>", "[:X]", 1, -1, 1)
			fString = replace(fString, "<img src=icon_smile_kisses.gif border=0 align=middle>", "[:x]", 1, -1, 1)
		end if
		if strAllowForumCode = "1" then
			if strIMGInPosts = "1" then
				fString = replace(fString, "<img src=""","[img]", 1, -1, 1)
				fString = replace(fString, "<img align=right src=""","[img=right]", 1, -1, 1)
				fString = replace(fString, "<img align=left src=""","[img=left]", 1, -1, 1)
				fString = replace(fString, """ border=0>","[/img]", 1, -1, 1)
				fString = replace(fString, """ id=right border=0>","[/img=right]", 1, -1, 1)
				fString = replace(fString, """ id=left border=0>","[/img=left]", 1, -1, 1)
				fString = replace(fString, "<img align=""right"" src=""","[img=right]", 1, -1, 1)
				fString = replace(fString, "<img align=""left"" src=""","[img=left]", 1, -1, 1)
				fString = replace(fString, """ border=""0"">","[/img]", 1, -1, 1)
				fString = replace(fString, """ id=""right"" border=""0"">","[/img=right]", 1, -1, 1)
				fString = replace(fString, """ id=""left"" border=""0"">","[/img=left]", 1, -1, 1)
			end if
		end if
	end if
	fString = Replace(fString, "'", "'")
	CleanCode = fString
end function
%>
