<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##       Huw Reddick and Richard Kinser
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

Response.Write	"<tr>" & vbNewLine & _
				"<td class=""formlabel"">Format Mode:&nbsp;</td>" & vbNewLine & _
				"<td class=""formvalue"">" & vbNewLine & _
				"<select name=""mode"" tabindex=""-1"" onChange=""thelp(this.options[this.selectedIndex].value)"">" & vbNewLine & _
				"<option selected value=""0"">Basic&nbsp;</option>" & vbNewLine & _
				"<option value=""1"">Help&nbsp;</option>" & vbNewLine & _
				"<option value=""2"">Prompt&nbsp;</option>" & vbNewLine & _
				"</select>" & vbNewLine & _
				"<a href=""JavaScript:openWindowHelp('pop_help.asp?mode=post#mode')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formlabel"" rowspan=""2"">Format:&nbsp;</td>" & vbNewLine & _
				"<td class=""formvalue"">" & vbNewLine & _
				"<a href=""Javascript:bold();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorBold,"Bold","align=""top""") & "</a>" & _
				"<a href=""Javascript:italicize();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorItalicize,"Italicized","align=""top""") & "</a>" & _
				"<a href=""Javascript:underline();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorUnderline,"Underline","align=""top""") & "</a>" & _
				"<a href=""Javascript:strike();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorStrike,"Strikethrough","align=""top""") & "</a>" & vbNewLine & _
				"<a href=""Javascript:left();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorLeft,"Align Left","align=""top""") & "</a>" & _
				"<a href=""Javascript:center();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorCenter,"Centered","align=""top""") & "</a>" & _
				"<a href=""Javascript:right();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorRight,"Align Right","align=""top""") & "</a>" & vbNewLine & _
				"<a href=""Javascript:hr();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorHR,"Horizontal Rule","align=""top""") & "</a>" & _
				"<a href=""Javascript:hyperlink();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorUrl,"Insert Hyperlink","align=""top""") & "</a>" & _
				"<a href=""Javascript:email();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorEmail,"Insert Email","align=""top""") & "</a>"

if strIMGInPosts = "1" then
	Response.Write	"<a href=""Javascript:image();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorImage,"Insert Image","align=""top""") & "</a>" & vbNewLine
end if

Response.Write	"<a href=""Javascript:showcode();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorCode,"Insert Code","align=""top""") & "</a>" & _
				"<a href=""Javascript:quote();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorQuote,"Insert Quote","align=""top""") & "</a>" & _
				"<a href=""Javascript:list();"" tabindex=""-1"">" & getCurrentIcon(strIconEditorList,"Insert List","align=""top""") & "</a>" & vbNewLine

if lcase(strIcons) = "1" and strShowSmiliesTable = "0" then
	Response.Write	"<a href=""JavaScript:openWindow6('pop_icon_legend.asp')"" tabindex=""-1"">" & getCurrentIcon(strIconEditorSmilie,"Insert Smilie","align=""top""") & "</a>" & vbNewLine
end if

Response.Write	"</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""formvalue"">" & vbNewLine & _
				"<select name=""Font"" tabindex=""-1"" onChange=""showfont(this.options[this.selectedIndex].value)"">" & vbNewLine & _
				"<option value="""" selected>Font</option>" & vbNewLine & _
				"<option value=""Andale Mono"">Andale Mono</option>" & vbNewLine & _
				"<option value=""Arial"">Arial</option>" & vbNewLine & _
				"<option value=""Arial Black"">Arial Black</option>" & vbNewLine & _
				"<option value=""Book Antiqua"">Book Antiqua</option>" & vbNewLine & _
				"<option value=""Century Gothic"">Century Gothic</option>" & vbNewLine & _
				"<option value=""Comic Sans MS"">Comic Sans MS</option>" & vbNewLine & _
				"<option value=""Courier New"">Courier New</option>" & vbNewLine & _
				"<option value=""Georgia"">Georgia</option>" & vbNewLine & _
				"<option value=""Impact"">Impact</option>" & vbNewLine & _
				"<option value=""Lucida Console"">Lucida Console</option>" & vbNewLine & _
				"<option value=""Script MT Bold"">Script MT Bold</option>" & vbNewLine & _
				"<option value=""Stencil"">Stencil</option>" & vbNewLine & _
				"<option value=""Tahoma"">Tahoma</option>" & vbNewLine & _
				"<option value=""Times New Roman"">Times New Roman</option>" & vbNewLine & _
				"<option value=""Trebuchet MS"">Trebuchet MS</option>" & vbNewLine & _
				"<option value=""Verdana"">Verdana</option>" & vbNewLine & _
				"</select>&nbsp;" & vbNewLine & _
				"<select name=""Size"" tabindex=""-1"" onChange=""showsize(this.options[this.selectedIndex].value)"">" & vbNewLine & _
				"<option value="""" selected>Size</option>" & vbNewLine & _
				"<option value=""1"">1 (x-small)</option>" & vbNewLine & _
				"<option value=""2"">2 (small)</option>" & vbNewLine & _
				"<option value=""3"">3 (medium)</option>" & vbNewLine & _
				"<option value=""4"">4 (large)</option>" & vbNewLine & _
				"<option value=""5"">5 (x-large)</option>" & vbNewLine & _
				"<option value=""6"">6 (xx-large)</option>" & vbNewLine & _
				"</select>&nbsp;" & vbNewLine & _
				"<select name=""Color"" tabindex=""-1"" onChange=""showcolor(this.options[this.selectedIndex].value)"">" & vbNewLine & _
				"<option value="""" selected>Color</option>" & vbNewLine & _
				"<option style=""color:black"" value=""black"">Black</option>" & vbNewLine & _
				"<option style=""color:red"" value=""red"">Red</option>" & vbNewLine & _
				"<option style=""color:yellow"" value=""yellow"">Yellow</option>" & vbNewLine & _
				"<option style=""color:pink"" value=""pink"">Pink</option>" & vbNewLine & _
				"<option style=""color:green"" value=""green"">Green</option>" & vbNewLine & _
				"<option style=""color:orange"" value=""orange"">Orange</option>" & vbNewLine & _
				"<option style=""color:purple"" value=""purple"">Purple</option>" & vbNewLine & _
				"<option style=""color:blue"" value=""blue"">Blue</option>" & vbNewLine & _
				"<option style=""color:beige"" value=""beige"">Beige</option>" & vbNewLine & _
				"<option style=""color:brown"" value=""brown"">Brown</option>" & vbNewLine & _
				"<option style=""color:teal"" value=""teal"">Teal</option>" & vbNewLine & _
				"<option style=""color:navy"" value=""navy"">Navy</option>" & vbNewLine & _
				"<option style=""color:maroon"" value=""maroon"">Maroon</option>" & vbNewLine & _
				"<option style=""color:limegreen"" value=""limegreen"">LimeGreen</option>" & vbNewLine & _
				"</select></td>" & vbNewLine & _
				"</tr>" & vbNewLine
%>
