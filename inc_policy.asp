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
'## MOD: Custom Policy v1.2 for Snitz Forums v3.4
'## Author: Michael Reisinger (OneWayMule)
'## File: inc_policy.asp
'##
'## Get the latest version of this MOD at
'## http://www.onewaymule.org/onewayscripts/
'#################################################################################

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Registration Rules and Policies Agreement</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

If strProhibitNewMembers <> "1" Then
	Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td>Registration Rules and Policies Agreement for " & strForumTitle & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & vbNewLine
	
	strsql = "SELECT CP_MODE, CP_CONTENT FROM " & strTablePrefix & "CUSTOM_POLICY WHERE CP_ID=1"
	Set rsp = my_conn.execute(strsql)
	strPolicyMode = rsp("CP_MODE")
	strPolicyContent = rsp("CP_CONTENT")
	strPolicyContent = formatStr(strPolicyContent)
	strPolicyContent = Replace(strPolicyContent,"[adminemail]","<a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">Contact the Administrator</a>")
	strPolicyContent = Replace(strPolicyContent,"[forumurl]","<a href=""" & strForumUrl & """>" & strForumUrl & "</a>")
	rsp.close
	set rsp = nothing
	
	If strPolicyMode = 1 Then
		Response.Write strPolicyContent
	Else
		Response.Write	"<p>If you agree to the terms and conditions stated below, " & _
						"press the &quot;Agree&quot; button. Otherwise, press &quot;Cancel&quot;.</p>" & vbNewLine & _
						"<p>In order to use these forums, users are required to " & _
						"provide a username, password and e-mail address. Neither the Administrators of " & _
						"these forums, or the Moderators participating, are responsible for the privacy " & _
						"practices of any user. Remember that all information that is disclosed in these " & _
						"areas becomes public information and you should exercise caution when deciding " & _
						"to share any of your personal information. Any user who finds material posted by " & _
						"another user objectionable is encouraged to contact us via e-mail. We are " & _
						"authorized by you to remove or modify any data submitted by you to these forums " & _
						"for any reason we feel constitutes a violation of our policies, whether stated, implied or not.</p>" & vbNewLine & _
						"<p>This site may contain links to other web sites and " & _
						"files. We have no control over the content and can not ensure it will not be offensive " & _
						"or objectionable.  We will, however, remove links to material that we feel is inappropriate as we become aware of them.</p>" & vbNewLine & _
						"<p>These forums give users two options for changing and " & _
						"modifying information that they provide in their profile: " & _
						"<ol>" & _
						"<li>Users can login with their username and password to " & _
						"change any information in their profile.</li>" & _
						"<li>In case of lost password, users can send an e-mail to " & _
						"<a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">the Administrator</a>.</li>" & _
						"</ol>" & _
						"</p>" & vbNewLine & _
						"<p>Cookies must be turned on in your browser to participate " & _
						"as a user in these forums. Cookies are used here to hold your username and " & _
						"password and viewing options, allowing you to login.</p>" & vbNewLine & _
						"<p>By pressing the &quot;Agree&quot; button, you agree that you, the user, are "
		
		If strMinAge > 0 Then
			Response.Write strMinAge
		Else
			Response.Write "13"
		End If
		
		Response.Write	" years of age or over. You are fully responsible for any information " & _
						"or file supplied by this user. You also agree that you will not post any " & _
						"copyrighted material that is not owned by yourself or the owners of these " & _
						"forums. In your use of these forums, you agree that you will not post any " & _
						"information which is vulgar, harassing, hateful, threatening, invading of others " & _
						"privacy, sexually oriented, or violates any laws.</p>" & vbNewLine & _
						"<p>If you do agree with the rules and policies stated in " & _
						"this agreement, and meet the criteria stated herein, proceed to press the " & _
						"&quot;Agree&quot; button below, otherwise press &quot;Cancel&quot;.</p>" & vbNewLine
	End If
	
	Response.Write	"<div class=""options"">" & vbNewLine & _
					"<form action=""register.asp?mode=Register"" id=""form1"" method=""post"" name=""form1"" style=""display:inline;"">" & vbNewLine & _
					"<input name=""Refer"" type=""hidden"" value=""" & strReferer & """>" & vbNewLine & _
					"<input name=""policy_accept"" type=""hidden"" value=""true"">" & vbNewLine & _
					"<button name=""Submit"" type=""Submit"">Agree</button>&nbsp;" & vbNewLine & _
					"</form>" & vbNewLine & _
					"<form action=""JavaScript:history.go(-1)"" id=""form2"" method=""post"" name=""form2"" style=""display:inline;"">" & vbNewLine & _
					"<button name=""Submit"" type=""Submit"">Cancel</button>" & vbNewLine & _
					"</div>" & vbNewLine & _
					"</form>" & vbNewLine
	
	If strPolicyMode <> 1 Then
		Response.Write	"</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td class=""options"">" & vbNewLine & _
						"<p>If you have any questions about this privacy statement " & _
						"or the use of these forums, you can " & _
						"<a href=""contact.asp""" & dWStatus("Contact the Administrator") & " tabindex=""-1"">" & _
						"contact the forum Administrator</a>.</p>" & vbNewLine
	End If
	
	Response.Write	"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine
Else
	Response.Write	"<div class=""warning"" style=""width:50%;"">" & vbNewLine & _
					"<p>Sorry, we are not accepting any new Members at this time.</p>" & vbNewLine & _
					"<meta http-equiv=""Refresh"" content=""5; URL=default.asp"">" & vbNewLine & _
					"<p><a href=""default.asp"">Back To Forum</a></p>" & vbNewLine & _
					"</div>" & vbNewLine
End If

WriteFooter
Response.End
%>
