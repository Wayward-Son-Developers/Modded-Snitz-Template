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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Ranking&nbsp;Configuration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Request.Form("strRankAdmin") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Administrator Name</li>"
	end if
 	if Request.Form("strRankMod") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Moderator Name</li>"
	end if
	if Request.Form("strRankLevel0") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Starting Member Name</li>"
	end if
	if Request.Form("strRankLevel1") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 1 Name</li>"
	end if
	if Request.Form("strRankLevel2") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 2 Name</li>"
	end if
	if Request.Form("strRankLevel3") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 3 Name</li>"
	end if
	if Request.Form("strRankLevel4") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 4 Name</li>"
	end if
	if Request.Form("strRankLevel5") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 5 Name</li>"
	end if
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel2")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 2</li>"
	end if
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel3")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 3</li>"
	end if
	if cLng(Request.Form("intRankLevel2")) > cLng(Request.Form("intRankLevel3")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 2 can not be higher than 3</li>"
	end if
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel4")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 4</li>"
	end if
	if cLng(Request.Form("intRankLevel2")) > cLng(Request.Form("intRankLevel4")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 2 can not be higher than 4</li>"
	end if
	if cLng(Request.Form("intRankLevel3")) > cLng(Request.Form("intRankLevel4")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 3 can not be higher than 4</li>"
	end if
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel5")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 5</li>"
	end if
	if cLng(Request.Form("intRankLevel2")) > cLng(Request.Form("intRankLevel5")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 2 can not be higher than 5</li>"
	end if
	if cLng(Request.Form("intRankLevel3")) > cLng(Request.Form("intRankLevel5")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 3 can not be higher than 5</li>"
	end if
	if cLng(Request.Form("intRankLevel4")) > cLng(Request.Form("intRankLevel5")) then 
		Err_Msg = Err_Msg & "<li>Rank Level 4 can not be higher than 5</li>"
	end if

	if Err_Msg = "" then
		for each key in Request.Form 
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next

		Application(strCookieURL & "ConfigLoaded") = ""
		Call OkMessage("Configuration Posted!","admin_home.asp","Back To Admin Home")
	else
		Call FailMessage(Err_Msg,True)
	end if
else
	arrStarColors = ("gold|silver|bronze|orange|red|purple|blue|cyan|green")
	arrIconStarColors = array(strIconStarGold,strIconStarSilver,strIconStarBronze,strIconStarOrange,strIconStarRed,strIconStarPurple,strIconStarBlue,strIconStarCyan,strIconStarGreen)
	strStarColor = split(arrStarColors, "|")

	Response.Write	"<form action=""admin_config_ranks.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Ranking Configuration</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Show Ranking:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<select name=""strShowRank"">" & vbNewLine & _
					"<option value=""0""" & chkSelect(strShowRank,0) & ">None</option>" & vbNewLine & _
					"<option value=""1""" & chkSelect(strShowRank,1) & ">Rank Only</option>" & vbNewLine & _
					"<option value=""2""" & chkSelect(strShowRank,2) & ">Stars Only</option>" & vbNewLine & _
					"<option value=""3""" & chkSelect(strShowRank,3) & ">Rank and Stars</option>" & vbNewLine & _
					"</select>" & vbNewLine & _
					"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#ShowRank')"">" & getCurrentIcon(strIconSmileQuestion,"ShowRank","") & "</a></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Administrator</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Administrator Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankAdmin"" size=""30"" value=""" & chkExistElse(chkString(strRankAdmin,"edit"),"Administrator") & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Administrator)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColorAdmin"" value=""" & strStarColor(c) & """" & chkRadio(strRankColorAdmin,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Moderator</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Moderator Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankMod"" size=""30"" value=""" & chkExistElse(chkString(strRankMod,"edit"),"Moderator") & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Moderator)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColorMod"" value=""" & strStarColor(c) & """" & chkRadio(strRankColorMod,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Starting Member</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Starting Member Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankLevel0"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel0,"edit"),"Starting Member") & """>" & vbNewLine & _
					"Number:&nbsp;<input type=""text"" name=""intRankLevel0"" size=""5"" value=""0"" readonly>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Member who has less than Member Level 1 but more than Starting Member Level posts)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Member Level 1</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Member Level 1 Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankLevel1"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel1,"edit"),"New Member") & """>" & vbNewLine & _
					"Number:&nbsp;<input type=""text"" name=""intRankLevel1"" size=""5"" value=""" & chkExistElse(intRankLevel1,50) & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 1 and Member Level 2 posts)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColor1"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor1,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Member Level 2</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Member Level 2 Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankLevel2"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel2,"edit"),"Junior Member") & """>" & vbNewLine & _
					"Number:&nbsp;<input type=""text"" name=""intRankLevel2"" size=""5"" value=""" & chkExistElse(intRankLevel2,100) & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 2 and Member Level 3 posts)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColor2"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor2,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Member Level 3</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Member Level 3 Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankLevel3"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel3,"edit"),"Average Member") & """>" & vbNewLine & _
					"Number:&nbsp;<input type=""text"" name=""intRankLevel3"" size=""5"" value=""" & chkExistElse(intRankLevel3,500) & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 3 and Member Level 4 posts)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColor3"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor3,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Member Level 4</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Member Level 4 Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankLevel4"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel4,"edit"),"Senior Member") & """>" & vbNewLine & _
					"Number:&nbsp;<input type=""text"" name=""intRankLevel4"" size=""5"" value=""" & chkExistElse(intRankLevel4,1000) & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 4 and Member Level 5 posts)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColor4"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor4,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">Member Level 5</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Member Level 5 Name:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue""><input type=""text"" name=""strRankLevel5"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel5,"edit"),"Advanced Member") & """>" & vbNewLine & _
					"Number:&nbsp;<input type=""text"" name=""intRankLevel5"" size=""5"" value=""" & chkExistElse(intRankLevel5,2000) & """>" & vbNewLine & _
					getCurrentIcon(strIconSmileQuestion,"(Member who has more than Member Level 5 posts)","") & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Star Color:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine
	for c = 0 to ubound(strStarColor)
		Response.Write	"<input type=""radio"" class=""radio"" name=""strRankColor5"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor5,strStarColor(c),true) & ">" & getCurrentIcon(arrIconStarColors(c),"","") & vbNewLine
	next
	Response.Write	"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"RankColor","") & "</a>&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td colspan=""2"">&nbsp;</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><button type=""submit"" id=""submit1"" name=""submit1"">Submit New Config</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if 
WriteFooter
Response.End
%>
