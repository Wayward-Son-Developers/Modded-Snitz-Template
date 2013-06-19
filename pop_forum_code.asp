<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'## Huw Reddick and Richard Kinser
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
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA02111-1307, USA.
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
<%
Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td><a name=""format""></a>How to format text with Bold, Italic, Quote, etc...</td>" & vbNewLine & _
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
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine
WriteFooterShort
Response.End
%>
