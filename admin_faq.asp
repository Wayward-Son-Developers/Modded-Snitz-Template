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
'## File: admin_faq.asp
'##
'## Get the latest version of this MOD at
'## http://www.onewaymule.org/onewayscripts/
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<%
If Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
End If

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;F.A.Q.&nbsp;Administration</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine
 
strfaction = Request.Querystring("action")
strFID = CLng(Request.Querystring("id"))

Select Case strfaction
	Case "submitcat"
		'###   SUBMIT CATEGORY   ###
		If Request.Querystring("do") <> "yes" Then
			Response.Write	"<form method=""post"" action=""admin_faq.asp?action=submitcat&do=yes"" name=""PostTopic"">" & vbNewLine & _
							"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td colspan=""2"">Submit New F.A.Q. Category</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Category Title:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<input type=""text"" name=""title"" size=""30"" maxlength=""100"" value="""">" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Auth Type:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<select name=""authlevel"">" & vbNewLine & _
							"<option value=""4"" selected>Admins only</option>" & vbNewLine & _
							"<option value=""3"">Admins, Moderators</option>" & vbNewLine & _
							"<option value=""2"">All Members</option>" & vbNewLine & _
							"<option value=""0"">All Visitors</option>" & vbNewLine & _
							"</select>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""options"" colspan=""2"">" & vbNewLine & _
							"<button type=""submit"">Submit</button>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</form>" & vbNewLine  
		Else
			strFormTitle = ChkString(Request.Form("title"),"SQLString")
			intFormLevel = CLng(Request.Form("authlevel"))
			
			Err_Msg = ""
			
			If trim(strFormTitle) = "" Then
				Err_Msg = Err_Msg & "<li>You Must Enter a Title for the Category</li>"
			End If
			
			If Err_Msg = "" Then
				strsql = "INSERT INTO " & strTablePrefix & "FAQ_CATEGORY (FCAT_TITLE, FCAT_ORDER, FCAT_LEVEL)"
				strsql = strsql & " VALUES ('" & strFormTitle & "',1,'" & intFormLevel & "')"
				my_conn.Execute (strsql),,adCmdText + adExecuteNoRecords        
				
				Call OkMessage("New Category Submitted","admin_faq.asp","Return to F.A.Q. Administration")
			Else
				Call FailMessage(Err_Msg,True)
			End If     
		End If
		
	Case "editcat"
		'###   EDIT CATEGORY   ###
		If Request.Querystring("do") <> "yes" Then
			strsql = "SELECT FCAT_ID, FCAT_TITLE, FCAT_LEVEL FROM " & strTablePrefix & "FAQ_CATEGORY WHERE FCAT_ID=" & strFID
			Set crs = my_conn.execute(strsql)
			intFCatID = crs("FCAT_ID")
			strFCatTitle = crs("FCAT_TITLE")
			intFLevel = crs("FCAT_LEVEL")
			crs.Close
			Set crs = nothing
			Response.Write	"<form method=""post"" action=""admin_faq.asp?action=editcat&id=" & strFID & "&do=yes"" name=""PostTopic"">" & vbNewLine & _
							"<input type=""hidden"" name=""id"" value=""" & intFCatID & """>" & vbNewLine & _
							"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td colspan=""2"">Edit F.A.Q. Category</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Category Title:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<input type=""text"" name=""title"" maxlength=""100"" value=""" & ChkString(strFCatTitle,"display") & """>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Auth Type:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<select name=""authlevel"">" & vbNewLine & _
							"<option value=""4""" & chkSelect(4,intFLevel) & ">Admins only</option>" & vbNewLine & _
							"<option value=""3""" & chkSelect(3,intFLevel) & ">Admins, Moderators</option>" & vbNewLine & _
							"<option value=""2""" & chkSelect(2,intFLevel) & ">All Members</option>" & vbNewLine & _
							"<option value=""0""" & chkSelect(0,intFLevel) & ">All Visitors</option>" & vbNewLine & _
							"</select>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _                   
							"<tr>" & vbNewLine & _
							"<td class=""options"" colspan=""2""><button type=""submit"">Edit</button></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</form>" & vbNewLine
		Else
			strFormTitle = ChkString(Request.Form("title"),"SQLString")
			intFormLevel = CLng(Request.Form("authlevel"))
			
			Err_Msg = ""
			If trim(strFormTitle) = "" Then
				Err_Msg = Err_Msg & "<li>You Must Enter a Title for the Category</li>"
			End If
			
			If Err_Msg = "" Then
				strsql = "UPDATE " & strTablePrefix & "FAQ_CATEGORY SET FCAT_TITLE='" & strFormTitle & "', FCAT_LEVEL=" & intFormLevel & " WHERE FCAT_ID=" & strFID
				my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
				
				Call OkMessage("Category Updated","admin_faq.asp","Return to F.A.Q. Administration")
			Else
				Call FailMessage(Err_Msg,True)
			End If     
		End If
	
	Case "deletecat"
		'###   DELETE CATEGORY   ###
		strFID = Request.Querystring("id")
		
		If Request.Querystring("do") <> "yes" Then
			Response.Write	"<div class=""warning"" style=""width:50%;"">Are you sure you want to delete this Category and all F.A.Q.s in it?<br />" & vbNewLine & _
							"<a href=""admin_faq.asp?action=deletecat&id=" & strFID & "&do=yes"">Yes</a> | <a href=""admin_faq.asp"">No</a></div>" & vbNewLine
		Else
			strsql = "DELETE FROM " & strTablePrefix & "FAQ WHERE F_FAQ_CATEGORY=" & strFID
			my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
			
			strsql = "DELETE FROM " & strTablePrefix & "FAQ_CATEGORY WHERE FCAT_ID=" & strFID
			my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
			
			Call OkMessage("Category Deleted","admin_faq.asp","Return to F.A.Q. Administration")
		End If
		
	Case "submit"
		'###   SUBMIT F.A.Q.   ###
		If Request.Querystring("do") <> "yes" Then
			Response.Write	"<script language=""JavaScript"" type=""text/javascript"" src=""inc_code.js""></script>" & vbNewLine & _
							"<form method=""post"" action=""admin_faq.asp?action=submit&do=yes"" name=""PostTopic"" id=""PostTopic"">" & vbNewLine & _
							"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td colspan=""2"">Submit New F.A.Q.</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Question:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<input type=""text"" name=""title"" size=""50"" maxlength=""255"" value="""">" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Category:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine 
			Call GetCategories(strFID)
			Response.Write	"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
			
			If strAllowForumCode = "1" And strShowFormatButtons = "1" Then
				%><!-- #INCLUDE FILE="inc_post_buttons.asp" --><%                                                         
			End If 
			
			Response.Write	"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Answer:&nbsp;<br /><br />" & vbNewLine
			
			If strAllowForumCode = "1" Then 
				Response.Write  "* <a href=""JavaScript:openWindow3('pop_forum_code.asp')"">Forum Code</a> is ON<br />" & vbNewLine
			Else 
				Response.Write  "* Forum Code is OFF<br />" & vbNewLine
			End If
			
			If strAllowHTML = "1" Then 
				Response.Write  "* HTML is ON<br />" & vbNewLine
			Else
				Response.Write  "* HTML is OFF<br />" & vbNewLine
			End If
			
			Response.Write	"</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<textarea name=""Message"" cols=""50"" rows=""7"" onselect=""storeCaret(this);"" onclick=""storeCaret(this);"" onkeyup=""storeCaret(this);"" onchange=""storeCaret(this);""></textarea>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""options"" colspan=""2""><button type=""submit"">Submit</button></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</form>" & vbNewLine
		Else
			strFormTitle = ChkString(Request.Form("title"),"SQLString")
			strFormDescription = ChkString(Request.Form("Message"),"message")
			intFormCategory = CLng(Request.Form("cat"))
			
			Err_Msg = ""
			If trim(strFormTitle) = "" Then
				Err_Msg = Err_Msg & "<li>You Must Enter A Question</li>"
			End If
			
			If trim(strFormDescription) = "" Then
				Err_Msg = Err_Msg & "<li>You Must Enter An Answer</li>"
			End If
			
			If Err_Msg = "" Then   
				strsql = "INSERT INTO " & strTablePrefix & "FAQ (F_FAQ_QUESTION, F_FAQ_ANSWER, F_FAQ_ORDER, F_FAQ_CATEGORY, F_FAQ_TYPE)"
				strsql = strsql & " VALUES ('" & strFormTitle & "', '" & strFormDescription & "', 1, " & intFormCategory & ", 0)"
				my_conn.execute(strsql),,adCmdText + adExecuteNoRecords        
				
				Call OkMessage("F.A.Q. Submitted","admin_faq.asp","Return to F.A.Q. Administration")
			Else
				Call FailMessage(Err_Msg,True)
			End If     
		End If
	
	Case "edit"
		'###   EDIT F.A.Q.   ###
		If Request.Querystring("do") <> "yes" Then
			strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER,  F_FAQ_CATEGORY, F_FAQ_TYPE FROM " & strTablePrefix & "FAQ WHERE F_ID=" & strFID
			Set frs = my_conn.execute(strsql)             
			intFID = frs("F_ID")
			strFTitle = frs("F_FAQ_QUESTION")
			strFDescription = frs("F_FAQ_ANSWER")
			intFCat = frs("F_FAQ_CATEGORY")
			intFType = frs("F_FAQ_TYPE")
			frs.Close
			Set frs = nothing
			
			If intFType > 0 Then
				Response.Redirect("admin_faq.asp")
			End If
			Response.Write	"<script language=""JavaScript"" type=""text/javascript"" src=""inc_code.js""></script>" & vbNewLine & _
							"<form method=""post"" action=""admin_faq.asp?action=edit&id=" & intFID & "&do=yes"" name=""PostTopic"">" & vbNewLine & _
							"<input type=""hidden"" name=""id"" value=""" & intFID & """>" & vbNewLine & _
							"<table class=""admin"">" & vbNewLine & _
							"<tr class=""header"">" & vbNewLine & _
							"<td colspan=""2"">Edit F.A.Q.</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Question:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<input type=""text"" name=""title"" size=""50"" maxlength=""255"" value=""" & ChkString(strFTitle,"display") & """>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Category:&nbsp;</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine
			
							Call GetCategories(intFCat)
			
			Response.Write	"</td>" & vbNewLine & _
							"</tr>" & vbNewLine
			
			If strAllowForumCode = "1" And strShowFormatButtons = "1" Then
				%><!-- #INCLUDE FILE="inc_post_buttons.asp" --><%
			End If 
			
			Response.Write	"<tr>" & vbNewLine & _
							"<td class=""formlabel"">Answer:&nbsp;<br /><br />" & vbNewLine
			
			If strAllowForumCode = "1" Then 
				Response.Write  "* <a href=""JavaScript:openWindow3('pop_forum_code.asp')"">Forum Code</a> is ON<br />" & vbNewLine
			Else 
				Response.Write  "* Forum Code is OFF<br />" & vbNewLine
			End If
			
			If strAllowHTML = "1" Then 
				Response.Write  "* HTML is ON<br />" & vbNewLine
			Else                    
				Response.Write  "* HTML is OFF<br />" & vbNewLine
			End If
			
			Response.Write  "</td>" & vbNewLine & _
							"<td class=""formvalue"">" & vbNewLine & _
							"<textarea name=""Message"" cols=""50"" rows=""7"" onselect=""storeCaret(this);"" onclick=""storeCaret(this);"" onkeyup=""storeCaret(this);"" onchange=""storeCaret(this);"">" & CleanCode(strFDescription) & "</textarea>" & vbNewLine & _
							"</td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"<tr>" & vbNewLine & _
							"<td class=""options"" colspan=""2""><button type=""submit"">Edit</button></td>" & vbNewLine & _
							"</tr>" & vbNewLine & _
							"</table>" & vbNewLine & _
							"</form>" & vbNewLine
		Else
			strFormTitle = ChkString(Request.Form("title"),"SQLString")
			strFormDescription = ChkString(Request.Form("Message"),"message")
			intFormCategory = CLng(Request.Form("cat"))
			
			Err_Msg = ""
			If trim(strFormTitle) = "" Then
				Err_Msg = Err_Msg & "<li>You Must Enter A Question</li>"
			End If
			
			If trim(strFormDescription) = "" Then
				Err_Msg = Err_Msg & "<li>You Must Enter An Answer</li>"
			End If
			
			If Err_Msg = "" Then   
				strsql = "UPDATE " & strTablePrefix & "FAQ SET F_FAQ_QUESTION='" & strFormTitle & "', F_FAQ_ANSWER='" & strFormDescription & "', F_FAQ_CATEGORY=" & intFormCategory & " WHERE F_ID=" & strFID
				my_conn.Execute(strsql),,adCmdText + adExecuteNoRecords
				
				Call OkMessage("F.A.Q. Updated","admin_faq.asp","Return to F.A.Q. Administration")
			Else
				Call FailMessage(Err_Msg,True)
			End If
		End If
	
	Case "delete"
		'###   DELETE F.A.Q.   ###
		strFID = Request.Querystring("id")
		
		If Request.Querystring("do") <> "yes" Then
			Response.Write	"<div class=""warning"" style=""width:50%;"">Are you sure you want to delete this F.A.Q.?<br />" & vbNewLine & _
							"<a href=""admin_faq.asp?action=delete&id=" & strFID & "&do=yes"">Yes</a> | <a href=""admin_faq.asp"">No</a></div>" & vbNewLine
		Else
			strsql = "DELETE FROM " & strTablePrefix & "FAQ WHERE F_ID=" & strFID
			my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
			
			Call OkMessage("F.A.Q. Deleted","admin_faq.asp","Return to F.A.Q. Administration")
		End If
	
	Case Else
		'###   F.A.Q. ADMINISTRATION MAIN MENU   ###
		strsql = "SELECT FCAT_ID, FCAT_TITLE, FCAT_ORDER, FCAT_LEVEL FROM " & strTablePrefix & "FAQ_CATEGORY ORDER BY FCAT_ORDER ASC, FCAT_ID ASC"
		Set crs = my_conn.Execute(strsql)
		
		Response.Write	"<table class=""admin"" width=""100%"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td>" & vbNewLine & _
						"<a href=""Javascript:openWindow3('admin_faq_order.asp')"">" & getCurrentIcon(strIconSort,"Set the order of F.A.Q.s and Categories","align=""right""") & "</a>" & vbNewLine & _
						"<a href=""admin_faq.asp?action=submitcat"">" & getCurrentIcon(strIconFolderNewTopic,"Submit New Category","align=""right""") & "</a>" & vbNewLine & _
						"F.A.Q. Administration</td>" & vbNewLine & _
						"</tr>" & vbNewLine 
		
		If crs.Eof Then
			Response.Write	"<tr><td>No Categories found.</td></tr>" & vbNewLine
		Else
			Do Until crs.eof  
				Response.Write  "<tr class=""section"">" & vbNewLine & _                     
								"<td>" & vbNewLine & _
								"<a href=""admin_faq.asp?action=deletecat&id=" & crs("FCAT_ID") & """>" & getCurrentIcon(strIconTrashCan,"Delete Category","align=""right""") & "</a>" & vbNewLine & _
								"<a href=""admin_faq.asp?action=editcat&id=" & crs("FCAT_ID") & """>" & getCurrentIcon(strIconPencil,"Edit Category","align=""right""") & "</a>" & vbNewLine & _
								"<a href=""admin_faq.asp?action=submit&id=" & crs("FCAT_ID") & """>" & getCurrentIcon(strIconFolderNewTopic,"Submit New F.A.Q.","align=""right""") & "</a>" & vbNewLine & _
								"<a name=""faqcat" & crs("FCAT_ID") & """>" & ChkString(crs("FCAT_TITLE"),"display") & "</a>&nbsp;" & GetFAQLevel(crs("FCAT_LEVEL")) & vbNewLine & _
								"</td>" & vbNewLine & _
								"</tr>" & vbNewLine
				
				strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER, F_FAQ_ORDER, F_FAQ_TYPE FROM " & strTablePrefix & "FAQ WHERE F_FAQ_CATEGORY=" & crs("FCAT_ID") & " ORDER BY F_FAQ_ORDER ASC, F_FAQ_QUESTION ASC"
				Set frs = my_conn.execute(strsql)
				
				If frs.Eof Then
					Response.Write  "<tr><td>No F.A.Q.s found.</td></tr>" & vbNewLine
				Else
					Do Until frs.eof
						Response.Write  "<tr>" & vbNewLine & _
										"<td>" & vbNewLine & _
										"<a href=""admin_faq.asp?action=delete&id=" & frs("F_ID") & """>" & getCurrentIcon(strIconTrashCan,"Delete F.A.Q.","align=""right""") & "</a>" & vbNewLine
						
						If frs("F_FAQ_TYPE")=0 Then
							Response.Write  "<a href=""admin_faq.asp?action=edit&id=" & frs("F_ID") & """>" & getCurrentIcon(strIconPencil,"Edit F.A.Q.","align=""right""") & "</a>" & vbNewLine
						End If
						
						Response.Write	ChkString(frs("F_FAQ_QUESTION"),"display") & vbNewLine & _
										"</td>" & vbNewLine & _
										"</tr>" & vbNewLine
						frs.Movenext
					Loop
				End If
				
				frs.close
				Set frs = nothing
				crs.Movenext
			Loop
		End If
		Response.Write  "</table>" & vbNewLine
		crs.close
		Set crs = nothing
End Select

WriteFooter
Response.End

Function GetFAQLevel(lvl)
  Select Case lvl
    Case 4 GetFAQLevel = "(Admins Only)"
    Case 3 GetFAQLevel = "(Admins, Moderators)"
    Case 2 GetFAQLevel = "(All Members)"
    Case 0 GetFAQLevel = "(All Visitors)"
  End Select
End Function

Function GetCategories(cat_id)
  strsql = "SELECT FCAT_ID, FCAT_TITLE FROM " & strTablePrefix & "FAQ_CATEGORY ORDER BY FCAT_TITLE ASC"
  Set rsCat = my_conn.execute(strsql)
  Response.Write  "<select name=""cat"">"
  do while not rsCat.EOF
    Response.Write  "<option value=""" & rsCat("FCAT_ID") & """ "
    If cat_id = rsCat("FCAT_ID") Then
      Response.Write  "SELECTED "
    End If
    Response.Write  ">" & rsCat("FCAT_TITLE") & "</option>"
    rsCat.MoveNext
  Loop
  rsCat.Close
  Set rsCat = nothing
  Response.Write "</select>"
  set rsCategories = nothing
End Function
%>