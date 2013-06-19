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
'## MOD: F.A.Q. Administration v1.1 for Snitz Forums v3.4
'## Author: Michael Reisinger (OneWayMule)
'## File: faq.asp
'##
'## Get the latest version of this MOD at
'## http://www.onewaymule.org/onewayscripts/
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
Dim arrFaqCat, arrFaqs, bolCatFound, bolFaqFound, iFaqCat, rsFaq, rsFCat
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td>" & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;Frequently Asked Questions</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td>FAQ Table of Contents</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td>" & vbNewLine & _
				"<p><ul>" & vbNewLine

strSql = "SELECT FCAT_ID, FCAT_TITLE FROM " & strTablePrefix & "FAQ_CATEGORY WHERE FCAT_LEVEL<=" & mLev & " ORDER BY FCAT_ORDER ASC, FCAT_ID ASC"
Set rsFCat = my_conn.Execute(strSql)
If Not rsFCat.EOF Then
	arrFaqCat = rsFCat.GetRows()
	bolCatFound = True
Else
	bolCatFound = False
End If

rsFCat.Close
Set rsFCat = Nothing

If bolCatFound Then
	For iFaqCat = 0 to ubound(arrFaqCat, 2)
		intFaqCatID = arrFaqCat(0, iFaqCat)
		strFaqCatTitle = ChkString(arrFaqCat(1, iFaqCat),"display")
		
		Response.Write  "<li><a href=""#faqcat" & intFaqCatID & """>" & strFaqCatTitle & "</a>" & vbNewLine & _
						"<ul>" & vbNewLine
		
		strSql2 = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_TYPE FROM " & strTablePrefix & "FAQ WHERE F_FAQ_CATEGORY=" & intFaqCatID & " ORDER BY F_FAQ_ORDER ASC, F_ID ASC;"
		
		Set rsFaq = my_conn.Execute(strSql2)
		If Not rsFaq.BOF Then rsFaq.MoveFirst
		If Not rsFaq.EOF Then
			arrFaq = rsFaq.GetRows()
			bolFaqFound = True
		Else
			bolFaqFound = False
		End If
		rsFaq.Close
		Set rsFaq = Nothing
		
		If bolFaqFound Then
			For iFaq = 0 to ubound(arrFaq, 2)
				intFaqID = arrFaq(0, iFaq)
				strFaqTitle = ChkString(arrFaq(1, iFaq), "display")
				intFaqType = arrFaq(2, iFaq)
				allowedFaq = True
				If (intFaqType = 2) and (strIcons<>"1") Then allowedFaq = False
				If (intFaqType = 4) and (strAllowForumCode<>"1") Then allowedFaq = False
				If (intFaqType = 13) and (strBadWordFilter<>"1") Then allowedFaq = False
				If (intFaqType = 14) and (stremail<>"1") Then allowedFaq = False
				If (intFaqType = 15) and (stremail<>"1") Then allowedFaq = False
				If (intFaqType = 2) and (strModeration<>"1") Then allowedFaq = False
				If allowedFaq Then Response.Write  "<li><a href=""#faq" & intFaqID & """>" & strFaqTitle & "</a></li>" & vbNewLine
			Next
		Else
			Response.Write  "<li>No FAQ found.</li>" & vbNewLine
		End If
		Response.Write  "</ul></li>" & vbNewLine
		Erase arrFaq
	Next
	Erase arrFaqCat
Else
	Response.Write  "<li><span class=""spnMessageText"">No Categories found.</span></li>" & vbNewLine
End If
Response.Write  "</ul>" & vbNewLine & _
				"</p>" & vbNewLine

If strEmail = "1" Then Response.Write	"<p><a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">Can't find your answer here? Send us an e-mail.</a></p>" & vbNewLine

Response.Write	"</td>" & vbNewLine & _
				"</tr>" & vbNewLine

strSql = "SELECT FCAT_ID, FCAT_TITLE FROM " & strTablePrefix & "FAQ_CATEGORY WHERE FCAT_LEVEL<=" & mLev & " ORDER BY FCAT_ORDER ASC, FCAT_ID ASC"
Set rsFCat = my_conn.Execute(strSql)

If Not rsFCat.EOF Then
	arrFaqCat = rsFCat.GetRows()
	bolCatFound = True
Else
	bolCatFound = False
End If

rsFCat.Close
Set rsFCat = Nothing

If bolCatFound Then
	For iFaqCat = 0 to ubound(arrFaqCat, 2)
		intFaqCatID = arrFaqCat(0, iFaqCat)
		strFaqCatTitle = ChkString(arrFaqCat(1, iFaqCat),"display")
		Response.Write	"<tr class=""section"">" & vbNewLine & _
						"<td><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faqcat" & intFaqCatID & """></a>" & strFaqCatTitle & "</td>" & vbNewLine & _
						"</tr>" & vbNewLine
		
		strSql2 = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_TYPE, F_FAQ_ANSWER FROM " & strTablePrefix & "FAQ WHERE F_FAQ_CATEGORY=" & intFaqCatID & " ORDER BY F_FAQ_ORDER ASC, F_ID ASC;"
		Set rsFaq = my_conn.Execute(strSql2)
		
		If Not rsFaq.EOF Then
			arrFaq = rsFaq.GetRows()
			bolFaqFound = True
		Else
			bolFaqFound = False
		End If
		rsFaq.Close
		Set rsFaq = Nothing
		
		If bolFaqFound Then
			For iFaq = 0 to ubound(arrFaq, 2)
				intFaqID = arrFaq(0, iFaq)
				strFaqTitle = ChkString(arrFaq(1, iFaq), "display")
				intFaqType = arrFaq(2, iFaq)
				strFaqAnswer = formatStr(arrFaq(3, iFaq))
				If intFaqType > 0 Then
					Call DefaultFAQ(intFaqType,intFaqID)
				Else
					Response.Write  "<tr>" & vbNewLine & _
									"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intFaqID & """></a>" & strFaqTitle & "</td>" & vbNewLine & _
									"</tr>" & vbNewLine & _
									"<tr>" & vbNewLine & _
									"<td class=""faqans""><p>" & strFaqAnswer & "</p></td>" & vbNewLine & _
									"</tr>" & vbNewLine
				End If
			Next
		Else
			Response.Write	"<li>No FAQ found.</li>" & vbNewLine
		End If
		Response.Write	"</ul></li>" & vbNewLine
		Erase arrFaq
	Next
	Erase arrFaqCat
End If

Response.Write	"</table>" & vbNewLine
WriteFooter

Sub DefaultFAQ(intTheID, intTheOtherID)
	Select Case intTheID
		Case 1
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Registering</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans""><p>"
			If strProhibitNewMembers = "1" Then
				Response.Write	"The Administrator has turned off Registration for this forum. Only registered members are able to log in."
			ElseIf strRequireReg = "1" Then
				Response.Write	"Yes, registration is required. Registration is free and only takes a few minutes. The only required fields are your Username, which may be your real name or a nickname, a Password, and a valid e-mail address.<br /><br />"
			Else
				Response.Write	"Registration is not required to view current topics on the Forum; however, if you wish to post a new topic or reply to an existing topic registration is required. Registration is free and only takes a few minutes. The only required fields are your Username, which may be your real name or a nickname, and a valid e-mail address.<br /><br />"
			End If
			If strProhibitNewMembers = "0" Then
				Response.Write	"The information you provide during registration is not outsourced or used for any advertising by " & strForumTitle & ".<br /><br />If you believe someone is sending you advertisements as a result of the information you provided through your registration, please notify us immediately."
			End If
			Response.Write  "</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 2
			if (strIcons = "1") then
				strSmileCode = array("[:)]","[:D]","[8D]","[:I]","[:p]","[}:)]","[;)]","[:o)]","[B)]","[8]","[:(]","[8)]","[:0]","[:(!]","[xx(]","[|)]","[:X]","[^]","[V]","[?]")
				strSmileDesc = array("smile","big smile","cool","blush","tongue","evil","wink","clown","black eye","eightball","frown","shy","shocked","angry","dead","sleepy","kisses","approve","disapprove","question")
				strSmileName = array(strIconSmile,strIconSmileBig,strIconSmileCool,strIconSmileBlush,strIconSmileTongue,strIconSmileEvil,strIconSmileWink,strIconSmileClown,strIconSmileBlackeye,strIconSmile8ball,strIconSmileSad,strIconSmileShy,strIconSmileShock,strIconSmileAngry,strIconSmileDead,strIconSmileSleepy,strIconSmileKisses,strIconSmileApprove,strIconSmileDisapprove,strIconSmileQuestion)

				Response.Write	"<tr>" & vbNewLine & _
								"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Smilies</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""faqans"">" & vbNewLine & _
								"<p>" & vbNewLine & _
								"You've probably seen others use smilies before in e-mail messages or other bulletin " & vbNewLine & _
								"board posts. Smilies are keyboard characters used to convey an emotion, such as a smile " & vbNewLine & _
								getCurrentIcon(strIconSmile,"","hspace=""10"" align=""absmiddle""") & " or a frown " & vbNewLine & _
								getCurrentIcon(strIconSmileSad,"","hspace=""10"" align=""absmiddle""") & ". This bulletin board " & vbNewLine & _
								"automatically converts certain text to a graphical representation when it is " & vbNewLine & _
								"inserted between brackets [].&nbsp; Here are the smilies that are currently " & vbNewLine & _
								"supported by &ldquo;" & strForumTitle & "&rdquo;:<br />" & vbNewLine & _
								"<table class=""admin"">" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td>" & vbNewLine & _
								"<table class=""nb"">" & vbNewLine

				for sm = 0 to 9
					Response.Write  "<tr>" & vbNewLine & _
									"<td>" & getCurrentIcon(strSmileName(sm),"","hspace=""10"" align=""absmiddle""") & "</td>" & vbNewLine & _
									"<td>" & strSmileDesc(sm) & "</td>" & vbNewLine & _
									"<td>" & strSmileCode(sm) & "</td>" & vbNewLine & _
									"</tr>" & vbNewLine
				next

				Response.Write	"</table>" & vbNewLine & _
								"</td>" & vbNewLine & _
								"<td>" & vbNewLine & _
								"<table class=""nb"">" & vbNewLine

				for sm = 10 to 19
					Response.Write  "<tr>" & vbNewLine & _
									"<td>" & getCurrentIcon(strSmileName(sm),"","hspace=""10"" align=""absmiddle""") & "</td>" & vbNewLine & _
									"<td>" & strSmileDesc(sm) & "</td>" & vbNewLine & _
									"<td>" & strSmileCode(sm) & "</td>" & vbNewLine & _
									"</tr>" & vbNewLine
				next
				
				Response.Write	"</table>" & vbNewLine & _
								"</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"</table></p>" & vbNewLine & _
								"</td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if    
		Case 3
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Creating a Hyperlink in your message</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>You can easily add a hyperlink to your message. " & vbNewLine & _
							"All that you need to do is type the URL (<b>" & strForumURL & "</b>), and it will automatically be converted to a URL (<a href=""" & strForumURL & """ target=""_blank"">" & strForumURL & "</a>). " & vbNewLine & _
							"The trick here is to make sure you prefix your URL with the <b>http://</b>, <b>https://</b> or <b>file://</b></p>" & vbNewLine & _
							"<p>Another way to add hyperlinks is to use the <b>[url]</b>linkto<b>[/url]</b> tags.</p>" & vbNewLine & _
							"<div class=""callout"">" & vbNewLine & _
							"<p><i>This Example:</i><br />" & vbNewLine & _
							"<b>[url]</b>" & strForumURL & "<b>[/url]</b> takes you home!</p>" & vbNewLine & _
							"<p><i>Outputs This:</i><br />" & vbNewLine & _
							"<a href=""" & strForumURL & """>" & strForumURL & "</a> takes you home!</p>" & vbNewLine & _
							"</div></p>" & vbNewLine & _
							"<p>You can also add a mailto link to your message by typing in your e-mail address." & vbNewLine & _
							"<div class=""callout"">" & vbNewLine & _
							"<p><i>This Example:</i><br />" & vbNewLine & _
							"<b>example@example.com</b></p>" & vbNewLine & _
							"<p><i>Outputs this:</i><br />" & vbNewLine & _
							"<a href=""mailto:example@example.com"">example@example.com</a></p>" & vbNewLine & _
							"</div></p>" & vbNewLine & _
							"<p>If you use this tag: <b>[url=&quot;</b>linkto<b>&quot;]</b>description<b>[/url]</b> you can add a description to the link." & vbNewLine & _
							"<div class=""callout"">" & vbNewLine & _
							"<p><i>This Example:</i><br />" & vbNewLine & _
							"Take me to <b>[url=&quot;" & strForumURL & "&quot;]</b>" & strForumTitle & "<b>[/url]</b></p>" & vbNewLine & _
							"<p><i>Outputs This:</i><br />" & vbNewLine & _
							"Take me to <a href=""" & strForumURL & """>" & strForumTitle & "</a></p>" & vbNewLine & _
							"</div>" & vbNewLine & _
							"<div class=""callout"">" & vbNewLine & _
							"<p><i>This Example:</i><br />" & vbNewLine & _
							"If you have a question <b>[url=&quot;example@example.com&quot;]</b>E-Mail Me<b>[/url]</b></p>" & vbNewLine & _
							"<p><i>Outputs This:</i><br />" & vbNewLine & _
							"If you have a question <a href=""mailto:example@example.com"">E-Mail Me</a></p>" & vbNewLine & _
							"</div></p>" & vbNewLine
			
			if (strIMGInPosts = "1") then
				Response.Write	"<p>You can make clickable images by combining the <b>[url=""</b>linkto<b>""]</b>description<b>[/url]</b> and <b>[img]</b>image_url<b>[/img]</b> tags." & vbNewLine & _
								"<div class=""callout"">" & vbNewLine & _
								"<p><i>This Example:</i><br />" & vbNewLine & _
								"<b>[url=&quot;" & strForumURL & "&quot;][img]</b>" & strTitleImage & "<b>[/img][/url]</b></p>" & vbNewLine & _
								"<p><i>Outputs This:</i><br />" & vbNewLine & _
								"<a href=""" & strForumURL & """ target=""_blank"">" & getCurrentIcon(strTitleImage & "||","","") & "</a></p>" & vbNewLine & _
								"</div></p>" & vbNewLine
			end if
			
			Response.Write  "</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 4
			if strAllowForumCode = "1" then
				Response.Write	"<tr>" & vbNewLine & _
								"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>How to format text with Bold, Italic, Quote, etc...</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""faqans"">" & vbNewLine & _
								"<p>There are several Forum Codes you may use to change the appearance " & vbNewLine & _
								"of your text.&nbsp; Following is the list of codes currently available:</p>" & vbNewLine & _
								"<blockquote>" & vbNewLine & _
								"<p>For Trademarks, Registered Trademarks, and Copyrights, there are some special tags - [tm], [r], and [c] respectively." & vbNewLine & _
								"<div class=""callout"">Trademark<b>[tm]</b> = Trademark&trade;.<br />Registered Trademark<b>[r]</b> = Registered Trademark&reg;.<br />Copyright<b>[c]</b> =  Copyright&copy;.</div></p>" & vbNewLine & _
								"<p><b>Bold:</b> Enclose your text with [b] and [/b]." & vbNewLine & _
								"<div class=""callout"">This is <b>[b]</b>bold<b>[/b]</b> text. = This is <b>bold</b> text.</div></p>" & vbNewLine & _
								"<p><b>Italic:</b> Enclose your text with [i] and [/i]." & vbNewLine & _
								"<div class=""callout"">This is <b>[i]</b>italic<b>[/i]</b> text. = This is <i>italic</i> text.</div></p>" & vbNewLine & _
								"<p><b>Underline:</b> Enclose your text with [u] and [/u]." & vbNewLine & _
								"<div class=""callout"">This is <b>[u]</b>underline<b>[/u]</b> text. = This is <u>underline</u> text.</div></p>" & vbNewLine & _
								"<p><b>Monospaced:</b> Enclose your text with [tt] and [/tt]." & vbNewLine & _
								"<div class=""callout"">This is <b>[tt]</b>Monospaced<b>[/tt]</b> text. = This is <tt>Monospaced</tt> text.</div></p>" & vbNewLine & _
								"<p><b>Striking Text:</b> Enclose your text with [s] and [/s]." & vbNewLine & _
								"<div class=""callout"">There was a <b>[s]</b>mistake<b>[/s]</b> learning opportunity. = There was a <s>mistake</s> learning opportunity.</div></p>" & vbNewLine & _
								"<p><b>Superscripts:</b> Enclose your text with [sup] and [/sup]." & vbNewLine & _
								"<div class=""callout"">The formula is E = MC<b>[sup]</b>2<b>[/sup]</b>. = The formula is E = MC<sup>2</sup>.</div></p>" & vbNewLine & _
								"<p><b>Subscripts:</b> Enclose your text with [sub] and [/sub]." & vbNewLine & _
								"<div class=""callout"">Water is H<b>[sub]</b>2<b>[/sub]</b>O. = Water is H<sub>2</sub>O.</div></p>" & vbNewLine & _
								"<p><b>Aligning Text Left:</b> Enclose your text with [left] and [/left]." & vbNewLine & _
								"<div class=""callout""><b>[left]</b>This text is going left...<b>[/left]</b> = <div style=""text-align:left;"">This text is going left...</div></div></p>" & vbNewLine & _
								"<p><b>Aligning Text Center:</b> Enclose your text with [center] and [/center]." & vbNewLine & _
								"<div class=""callout""><b>[center]</b>This text is going center...<b>[/center]</b> = <div style=""text-align:center;"">This text is going center...</div></div></p>" & vbNewLine & _
								"<p><b>Aligning Text Right:</b> Enclose your text with [right] and [/right]." & vbNewLine & _
								"<div class=""callout""><b>[right]</b>This text is going right...<b>[/right]</b> = <div style=""text-align:right;"">This text is going right...</div></div></p>" & vbNewLine & _
								"<p><b>Horizontal Rule:</b> Place a horizontal line in your post with [hr]." & vbNewLine & _
								"<div class=""callout"">Text<b>[hr]</b>And More Text = <br /><br />Text<hr noshade size=""1"">And More Text</div></p>" & vbNewLine & _
								"<p><b>Call-outs:</b> Enclose your text with [callout] and [/callout]." & vbNewLine & _
								"<div class=""callout""><b>[callout]</b>Each example has been in a callout.<b>[/callout]</b> = <div class=""callout"">Each example has been in a callout.</div></div></p>" & vbNewLine & _
								"<p><b>Font Colors:</b> Enclose your text with [<i>fontcolor</i>] and [/<i>fontcolor</i>]<br />" & vbNewLine & _
								"<div class=""callout""><b>[red]</b>Text<b>[/red]</b> = <span style=""color:red;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[green]</b>Text<b>[/green]</b> = <span style=""color:green;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[blue]</b>Text<b>[/blue]</b> = <span style=""color:blue;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[white]</b>Text<b>[/white]</b> = <span style=""color:white;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[purple]</b>Text<b>[/purple]</b> = <span style=""color:purple;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[yellow]</b>Text<b>[/yellow]</b> = <span style=""color:yellow;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[violet]</b>Text<b>[/violet]</b> = <span style=""color:violet;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[brown]</b>Text<b>[/brown]</b> = <span style=""color:brown;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[black]</b>Text<b>[/black]</b> = <span style=""color:black;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[pink]</b>Text<b>[/pink]</b> = <span style=""color:pink;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[orange]</b>Text<b>[/orange]</b> = <span style=""color:orange;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[gold]</b>Text<b>[/gold]</b> = <span style=""color:gold;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[beige]</b>Text<b>[/beige]</b> = <span style=""color:beige;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[teal]</b>Text<b>[/teal]</b> = <span style=""color:teal;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[navy]</b>Text<b>[/navy]</b> = <span style=""color:navy;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[maroon]</b>Text<b>[/maroon]</b> = <span style=""color:maroon;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[limegreen]</b>Text<b>[/limegreen]</b> = <span style=""color:limegreen;"">Text</span></div>" & vbNewLine & _
								"</p>" & vbNewLine & _
								"<p><b>Headings:</b> Enclose your text with [h<i>number</i>] and [/h<i>n</i>] where <i>n</i> is 1-6.<br />" & vbNewLine & _
								"<div class=""callout""><b>[h1]</b>Text<b>[/h1]</b> = <h1>Text</h1></div>" & vbNewLine & _
								"<div class=""callout""><b>[h2]</b>Text<b>[/h2]</b> = <h2>Text</h2></div>" & vbNewLine & _
								"<div class=""callout""><b>[h3]</b>Text<b>[/h3]</b> = <h3>Text</h3></div>" & vbNewLine & _
								"<div class=""callout""><b>[h4]</b>Text<b>[/h4]</b> = <h4>Text</h4></div>" & vbNewLine & _
								"<div class=""callout""><b>[h5]</b>Text<b>[/h5]</b> = <h5>Text</h5></div>" & vbNewLine & _
								"<div class=""callout""><b>[h6]</b>Text<b>[/h6]</b> = <h6>Text</h6></div>" & vbNewLine & _
								"</p>" & vbNewLine & _
								"<p><b>Font Sizes:</b> Enclose your text with [size=<i>number</i>] and [/size=<i>n</i>] where <i>n</i> is 1-6.<br />" & vbNewLine & _
								"<div class=""callout""><b>[size=1]</b>Text<b>[/size=1]</b> = <span style=""font-size:x-small;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[size=2]</b>Text<b>[/size=2]</b> = <span style=""font-size:small;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[size=3]</b>Text<b>[/size=3]</b> = <span style=""font-size:medium;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[size=4]</b>Text<b>[/size=4]</b> = <span style=""font-size:large;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[size=5]</b>Text<b>[/size=5]</b> = <span style=""font-size:x-large;"">Text</span></div>" & vbNewLine & _
								"<div class=""callout""><b>[size=6]</b>Text<b>[/size=6]</b> = <span style=""font-size:xx-large;"">Text</span></div>" & vbNewLine & _
								"</p>" & vbNewLine & _
								"<p><b>Bulleted List:</b> <b>[list]</b> and <b>[/list]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>." & vbNewLine & _
								"<div class=""callout""><b>[list]</b><b>[*]</b>Item 1<b>[/*]</b><b>[*]</b>Item 2<b>[/*]</b><b>[*]</b>Item 3<b>[/*]</b><b>[/list]</b> = " & _
								"<p><ul><li>Item 1</li><li>Item 2</li><li>Item 3</li></ul></p></div></p>" & vbNewLine & _
								"<p><b>Ordered Alpha List:</b> <b>[list=a]</b> and <b>[/list=a]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>." & vbNewLine & _
								"<div class=""callout""><b>[list=a]</b><b>[*]</b>Item 1<b>[/*]</b><b>[*]</b>Item 2<b>[/*]</b><b>[*]</b>Item 3<b>[/*]</b><b>[/list=a]</b> = " & _
								"<p><ol class=""alpha""><li>Item 1</li><li>Item 2</li><li>Item 3</li></ol></p></div></p>" & vbNewLine & _
								"<p><b>Ordered Number List:</b> <b>[list=1]</b> and <b>[/list=1]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>." & vbNewLine & _
								"<div class=""callout""><b>[list=1]</b><b>[*]</b>Item 1<b>[/*]</b><b>[*]</b>Item 2<b>[/*]</b><b>[*]</b>Item 3<b>[/*]</b><b>[/list=1]</b> = " & _
								"<p><ol class=""decimal""><li>Item 1</li><li>Item 2</li><li>Item 3</li></ol></p></div></p>" & vbNewLine & _
								"<p><b>Code:</b> Enclose your text with <b>[code]</b> and <b>[/code]</b>." & vbNewLine & _
								"<div class=""callout""><p><b>[code]</b>10 PRINT ""I ROCK AT BASIC!""<br />20 GOTO 10<br /><b>[/code]</b></p> = <p>" & _
								"<pre id=""code"">10 PRINT ""I ROCK AT BASIC!""" & vbNewLine & "20 GOTO 10" & vbNewLine & "</pre id=""code""></p></div></p>" & vbNewLine & _
								"<p><b>Scroll Code:</b> Enclose your text with <b>[scrollcode]</b> and <b>[/scrollcode]</b>. (Useful if you are posting a large chunk of code)" & vbNewLine & _
								"<div class=""callout""><p><b>[scrollcode]</b>10 PRINT ""I ROCK AT BASIC!""<br />20 GOTO 10<br /><b>[/scrollcode]</b></p> = <p>" & _
								"<div id=""scrollcode"" class=""scrollcode""><pre>10 PRINT ""I ROCK AT BASIC!""" & vbNewLine & "20 GOTO 10" & vbNewLine & "</pre></div id=""scrollcode""></p></div></p>" & vbNewLine & _
								"<p><b>Quote:</b> Enclose your text with <b>[quote]</b> and <b>[/quote]</b>." & vbNewLine & _
								"<div class=""callout""><b>[quote]</b>To be or not to be - that is the question.<b>[/quote]</b> = " & _
								"<blockquote class=""quote"">quote:<hr height=""1"" noshade id=""quote"">To be or not to be - that is the question.<hr height=""1"" noshade id=""quote""></blockquote id=""quote""></div></p>" & vbNewLine
				
				if (strIMGInPosts = "1") then
					Response.Write	"<p><b>Images:</b> Enclose the address with one of the following:<ul><li><b>[img]</b> and <b>[/img]</b></li>" & vbNewLine & _
									"<li><b>[img=right]</b> and <b>[/img=right]</b></li>" & vbNewLine & _
									"<li><b>[img=left]</b> and <b>[/img=left]</b></li></ul></p>" & vbNewLine
				end if
				
				Response.Write	"</blockquote></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
		Case 5
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Moderators</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>Moderators control individual forums. They may edit, delete, or prune any posts in their forums." 
			if (strShowModerators = "1") then
				Response.Write	" If you have a question about a particular forum, you should direct it to your forum moderator."
			end if
			Response.Write  "</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 6
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Cookies</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>These Forums use cookies to store the following information: the last time you logged in, your Username and your Encrypted Password. These cookies are stored on your hard drive. Cookies are not used to track your movement or perform any function other than to enhance your use of these forums."
			if (strNoCookies = "0") then
				Response.Write	" If you have not enabled cookies in your browser, many of these time-saving features will not work properly. <b>Also, you need to have cookies enabled if you want to enter a private forum or post a topic/reply.</b>"
			end if
			Response.Write  "</p>" & vbNewLine & _
							"<p>You may delete all cookies set by these forums in selecting the &quot;logout&quot; button at the top of any page.</p>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 7
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Active Topics</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>Active Topics are tracked by cookies. When you click on the &quot;active topics&quot; link, a page is generated listing all topics that have been posted since your last visit to these forums (or approximately 20 minutes).</p>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 8
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Editing Your Posts</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>You may edit or delete your own posts at any time. Just go to the topic where the post to be edited or deleted is located and you will see a edit or delete icon (" & getCurrentIcon(strIconEditTopic,"Edit","align=""absmiddle"" hspace=""6""") & getCurrentIcon(strIconDeleteReply,"Delete","align=""absmiddle"" hspace=""6""") & ") on the line that begins &quot;posted on...&quot; Click on this icon to edit or delete the post. No one else can edit your post, except for the forum Moderator or the forum Administrator. "
			if (strEditedByDate = "1") then
				Response.Write	"A note is generated at the bottom of each edited post displaying when and by whom the post was edited."
			end if
			Response.Write  "</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 9
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Attaching Files</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>For security reasons, you may not attach files to any posts. However, you may cut and paste text into your post.</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 10
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Searching For Specific Posts</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>You may search for specific posts based on a word or words found in the posts, user name, date, and particular forum(s). Simply click on the &quot;search&quot; link at the top of most pages.</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 11
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Editing Your Profile</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans""><p>You may easily change any information stored in your registration profile by using the ""profile"" link located near the top of each page. Simply identify yourself by typing your Username and Password and all of your profile information will appear on screen. You may edit any information (except your Username).</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 12
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Signatures</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans""><p>You may attach signatures to the end of your posts when you post either a New Topic or Reply. Your signature is editable by clicking on &quot;profile&quot; at the top of any forum page and entering your Username and Password.</p>" & vbNewLine & _
							"<p>NOTE: HTML can't be used in Signatures.</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 13
			if (strBadWordFilter = "1") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Censoring Posts</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""faqans"">" & vbNewLine & _
								"<p>The Forum does censor certain words that may be posted; however, this censoring is not an exact science, and is being done based on the words that are being screened, so certain words may be censored out of context. By default, words that are censored are replaced with asterisks.</p></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if    
		Case 14
			if (stremail = "1") then
				Response.Write	"<tr>" & vbNewLine & _
								"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Lost Password</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""faqans"">" & vbNewLine & _
								"<p>Changing a lost password is simple, assuming that e-mail features are turned on for this forum. All of the pages that require you to identify yourself with your Username and Password carry a &quot;lost Password&quot; link that you can use to have a code e-mailed instantly to your e-mail address of record that will allow you to create a new password. Because of the Encryption that we use for your password, we cannot tell you what your password is.</p></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
		Case 15
			if (stremail = "1") then
				Response.Write  "<tr>" & vbNewLine & _
								"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Can I be notified by e-mail when there are new posts?</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""faqans"">" & vbNewLine & _
								"<p>Yes, the <b>Subscription</b> feature allows you to subscribe to the entire Board, individual Categories, Forums and/or Topics, depending on what the administrator of this site allows. You will receive an e-mail notifying you of a post that has been made to the Category/Forum/Topic that you have subscribed to. There are four levels of subscription:" & vbNewLine & _
								"<ul>" & vbNewLine & _
								"<li><b>Board Wide Subscription</b><br />" & vbNewLine & _
								"If you can subscribe to an entire Board, you'll get a notification for any posts made within all the forums inside that board.</li>" & vbNewLine & _
								"<li><b>Category Wide Subscription</b><br />" & vbNewLine & _
								"You can subscribe to an entire Category, which will notify you if there was any posts made within any topic, within any forum, within that Category.</li>" & vbNewLine & _
								"<li><b>Forum Wide Subscription</b><br />" & vbNewLine & _
								"If you don't want to subscribe to an entire Category, you can subscribe to a single forum. This will notify you of any posts made within any topic, within that forum.</li>" & vbNewLine & _
								"<li><b>Topic Wide Subscription</b><br />" & vbNewLine & _
								"More conveniently, you can subscribe to just an individual topic. You will be notified of any post made within that topic.</li>" & vbNewLine & _
								"</ul>" & vbNewLine & _
								"Each level of subscription is optional. The administrator can turn <b>On/Off</b> each level of subscription for each Category/Forum/Topic. " & vbNewLine & _
								"To Subscribe or Unsubscribe from any level of subscription, you can use the ""My Subscriptions"" link, located near the top of each page to manage your subscriptions. Or you can click on the subscribe/unsubscribe icons (" & getCurrentIcon(strIconSubscribe,"Subscribe","align=""absmiddle""") & "&nbsp;" & getCurrentIcon(strIconUnsubscribe,"UnSubscribe","align=""absmiddle""") & ") for that Category/Forum/Topic you want to subscribe/unsubscribe to/from.</p></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
		Case 16
			if (strModeration = "1") then
				Response.Write  "<tr>" & vbNewLine & _
								"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>What does it mean if a forum has Moderation enabled?</td>" & vbNewLine & _
								"</tr>" & vbNewLine & _
								"<tr>" & vbNewLine & _
								"<td class=""faqans"">" & vbNewLine & _
								"<p><b>Moderation:</b> This feature allows the Administrator or the Moderator to ""<b>Approve</b>"", ""<b>Hold</b>"" or ""<b>Delete</b>"" a users post before it is shown to the public." & vbNewLine & _
								"<ul><li><b>Approve:</b> Only the administrators or the moderators will be able to approve a post made to a moderated forum. When the post is approved, it will be made viewable to the public.</li>" & vbNewLine & _
								"<li><b>Hold:</b> When a user posts a message to a moderated forum, the message is automatically put on hold until a moderator or an administrator approves of the post. No one will be able to view the post while it is put on hold. " & vbNewLine & _
								"(<i>NOTE: Authors of the post will be able to edit their post during this mode.</i>)</li>" & vbNewLine & _
								"<li><b>Delete:</b> If the administrator or moderator chooses this option, the post will be deleted and an e-mail will be sent to the poster of the message, informing them that their post was not approved. The administrator/moderator will be able to give their reason for not approving the post in the e-mail.</li>" & vbNewLine & _
								"</ul></p></td>" & vbNewLine & _
								"</tr>" & vbNewLine
			end if
		Case 17
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>What is COPPA?</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>The Children's Online Privacy Protection Act and Rule apply to individually identifiable information about a child that is collected online, such as full name, home address, e-mail address, telephone number or any other information that would allow someone to identify or contact the child. The Act and Rule also cover other types of information -- for example, hobbies, interests and information collected through cookies or other types of tracking mechanisms -- when they are tied to individually identifiable information. More information can be found <a href=""http://www.ftc.gov/bcp/conline/pubs/buspubs/coppa.htm"" title=""What is COPPA?"">here</a>.</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 18
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>Getting Your Own Forum</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>The most recent version of this Snitz Forum can be downloaded at <a href=""http://forum.snitz.com/"" target=""_blank"" title=""Link to Snitz Forums 2000 Homepage!"">this Internet web site</a>.</p>" & vbNewLine & _
							"<p>NOTE: The software is highly configurable, and the baseline Snitz Forum may not have all the features this forum does.</p></td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 19
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>What is RSS?</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p>As per Wikipedia.org: RSS (which, in its most recent format, stands for ""Really Simple Syndication"") is a family of web " & _
							"feed formats used to publish frequently updated content such as blog entries, news headlines or podcasts. An RSS document," & _
							"which is called a ""feed"", ""web feed"", or ""channel"", contains either a summary of content from an associated web site or" & _
							"the full text. RSS makes it possible for people to keep up with their favorite web sites in an automated manner that's " & _
							"easier than checking them manually.</p>" & _
							"<p>RSS content can be read using software called a &ldquo;feed reader&rdquo; or an &ldquo;aggregator.&rdquo; The user subscribes to a feed by " & _
							"entering the feed's link into the reader or by clicking an RSS icon in a browser that initiates the subscription process. " & _
							"The reader checks the user's subscribed feeds regularly for new content, downloading any updates that it finds.</p>" & _
							"<p>For a good primer on more than you ever wanted to know on RSS, <a href=""http://en.wikipedia.org/wiki/RSS"" target=""_blank"">read " & _
							"the entire Wikipedia.org article</a>.</p>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		Case 20
			Response.Write  "<tr>" & vbNewLine & _
							"<td class=""faq""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a><a name=""faq" & intTheOtherID & """></a>How do I access an RSS Feed?</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""faqans"">" & vbNewLine & _
							"<p><a href=""http://en.wikipedia.org/wiki/Aggregator"" target=""_blank"">RSS Aggregators</a> (also called Readers) will download and " & _
							"display RSS feeds for you. A number of free and commercial News Aggregators are available. Many aggregators are separate, " & _
							"&ldquo;stand-alone&rdquo; programs; other services will let you add RSS feeds to a Web page.</p>" & _
							"<p>To view one of the " & strForumTitle & "'s feeds in your RSS Aggregator:" & _
							"<ol>" & _
							"<li>Copy the URL/shortcut that corresponds to the topic that interests you.</li>" & _
							"<li>Paste the URL into your reader.</li></ol></p>" & _
							"<p>There are 3 feeds you can subscribe to:</p>" & _
							"<table class=""admin"" width=""80%"">" & _
							"<tr>" & _
							"<td>" & _
							"<p><a href=""rss.asp"" target=""_blank"">" & getCurrentIcon(strIconRSSPublicFeed,"Public RSS Feed"," border=""0""") & "</a></p>" & _
							"<p><a href=""rss.asp"" target=""_blank"">rss.asp</a></p><p>This feed will give anyone a list of the 20 most recent posts in the forum.</p>" & _
							"</td>" & _
							"</tr>" & _
							"<tr>" & _
							"<td>" & _
							"<p><a href=""rssfeed.asp"" target=""_blank"">" & getCurrentIcon(strIconRSSPublicReader,"Web Reader for Public RSS Feed"," border=""0""") & "</a></p>" & _
							"<p><a href=""rssfeed.asp"" target=""_blank"">rssfeed.asp</a></p><p>This is a web-based feed reader designed " & _
							"soley to display the public feed in a more friendly format. Some sidebar gadgets and/or other programs allow " & _
							"you to view web snippets and/or web pages. You can use this link if you have a program that will do this, but " & _
							"doesn't have an RSS Reader.</p>" & _
							"</td>" & _
							"</tr>"
			if MemberID > 0 then
				Response.Write	"<tr>" & _
								"<td>" & _
								"<p><a href=""" & strRssURL & """ target=""_blank"">" & getCurrentIcon(strIconRSSPrivateFeed,"Personal RSS Feed"," border=""0""") & "</a></p>" & _
								"<p><a href=""" & strRssURL & """ target=""_blank"">" & strRssURL & "</a></p>" & _
								"<p>This is a private feed. It passes your Member ID and a checksum to identify to the feed who you are. This " & _
								"adds information to the feed from any hidden topics you may have access to. <span class=""hlf"">Do NOT " & _
								"share this or subscribe to it from a public account.</span></p>" & _
								"</td>" & _
								"</tr>" & vbNewLine
			Else
				Response.Write	"<tr>" & _
								"<td>" & _
								"<p>" & getCurrentIcon(strIconRSSPrivateFeed,"Personal RSS Feed"," border=""0""") & "</p>" & _
								"<p>This is a personal feed available to registered members. You must log in to see the link.</p>" & _
								"</td>" & _
								"</tr>" & vbNewLine
			end if
			Response.Write	"</table>" & vbNewLine & _
							"</p>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
	End Select
End Sub
Response.End
%>