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
<%
strSmileCode = array("[:)]","[:D]","[8D]","[:I]","[:p]","[}:)]","[;)]","[:o)]","[B)]","[8]","[:(]","[8)]","[:0]","[:(!]","[xx(]","[|)]","[:X]","[^]","[V]","[?]")
strSmileDesc = array("smile","big smile","cool","blush","tongue","evil","wink","clown","black eye","eightball","frown","shy","shocked","angry","dead","sleepy","kisses","approve","disapprove","question")
strSmileName = array(strIconSmile,strIconSmileBig,strIconSmileCool,strIconSmileBlush,strIconSmileTongue,strIconSmileEvil,strIconSmileWink,strIconSmileClown,strIconSmileBlackeye,strIconSmile8ball,strIconSmileSad,strIconSmileShy,strIconSmileShock,strIconSmileAngry,strIconSmileDead,strIconSmileSleepy,strIconSmileKisses,strIconSmileApprove,strIconSmileDisapprove,strIconSmileQuestion)

Response.Write	"<script language=""Javascript"" type=""text/javascript"">" & vbNewLine & _
				"      <!-- hide" & vbNewLine & _
				"      function insertsmilie(smilieface) {" & vbNewLine & _
				"      		if (window.opener.document.PostTopic.Message.createTextRange && window.opener.document.PostTopic.Message.caretPos) {" & vbNewLine & _
				"       		var caretPos = window.opener.document.PostTopic.Message.caretPos;" & vbNewLine & _
				"               	caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? smilieface + ' ' : smilieface;" & vbNewLine & _
				"               	window.opener.document.PostTopic.Message.focus();" & vbNewLine & _
				"       	} else {" & vbNewLine & _
				"      			window.opener.document.PostTopic.Message.value+=smilieface;" & vbNewLine & _
				"                      	window.opener.document.PostTopic.Message.focus();" & vbNewLine & _
				"      		}" & vbNewLine & _
				"      }" & vbNewLine & _
				"      // -->" & vbNewLine & _
				"</script>" & vbNewLine & _
				"<table class=""content"" width=""100%"">" & vbNewLine & _
				"<tr class=""header"">" & vbNewLine & _
				"<td><a name=""smilies""></a>Smilies</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td>" & vbNewLine & _
				"<p>You've probably seen others use smilies before in e-mail messages or other bulletin " & vbNewLine & _
				"board posts. Smilies are keyboard characters used to convey an emotion, such as a smile " & vbNewLine & _
				getCurrentIcon(strIconSmile,"","hspace=""10"" align=""absmiddle""") & " or a frown " & vbNewLine & _
				getCurrentIcon(strIconSmileSad,"","hspace=""10"" align=""absmiddle""") & ". This bulletin board " & vbNewLine & _
				"automatically converts certain text to a graphical representation when it is " & vbNewLine & _
				"inserted between brackets [].&nbsp; Here are the smilies that are currently " & vbNewLine & _
				"supported by &ldquo;" & strForumTitle & "&rdquo;:<br /><br />" & vbNewLine & _
				"<table class=""admin"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td>" & vbNewLine & _
				"<table class=""nb"">" & vbNewLine

for sm = 0 to 9
	Response.Write  "<tr>" & vbNewLine & _
					"<td><a href=""Javascript:insertsmilie('" & strSmileCode(sm) & "');"">" & getCurrentIcon(strSmileName(sm),"","hspace=""10"" align=""absmiddle""") & "</a></td>" & vbNewLine & _
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
					"<td><a href=""Javascript:insertsmilie('" & strSmileCode(sm) & "');"">" & getCurrentIcon(strSmileName(sm),"","hspace=""10"" align=""absmiddle""") & "</a></td>" & vbNewLine & _
					"<td>" & strSmileDesc(sm) & "</td>" & vbNewLine & _
					"<td>" & strSmileCode(sm) & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine
next

Response.Write	"</table>" & vbNewLine & _
				"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table></p>" & vbNewLine & _
				"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine
WriteFooterShort
Response.End
%>
