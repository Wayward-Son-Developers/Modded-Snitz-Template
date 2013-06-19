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
'## File: admin_faq_order.asp
'##
'## Get the latest version of this MOD at
'## http://www.onewaymule.org/onewayscripts/
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%

if Session(strCookieURL & "Approval") <> strAdminCode then
	Response.Redirect "admin_login_short.asp?target=admin_faq_order.asp"
end if

If Request.Form("Method_Type") = "Write_Configuration" Then 
	If Request.Form("NumberCategories") <> "" Then
		i = 1
		Do Until i > cLng(Request.Form("NumberCategories"))
			SelectName = Request.Form("SortCategory" & i)
			If isNull(SelectName) Then SelectName = cLng(Request.Form("NumberCategories"))
			
			SelectID = Request.Form("SortCatID" & i)
			NumberFAQs = Request.Form("NumberFAQs" & SelectID)
			strsql = "UPDATE " & strTablePrefix & "FAQ_CATEGORY SET FCAT_ORDER=" & SelectName & " WHERE FCAT_ID=" & SelectID
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			
			If NumberFAQs <> "" Then
				j = 1
				Do Until j > cLng(Request.Form("NumberFAQs" & SelectID))
					SelectNamec = Request.Form("SortCat" & i & "SortFAQ" & j)
					
					If isNull(SelectNamec) Then SelectNamec = cLng(Request.Form("NumberFAQs" & SelectID))
					
					SelectIDc = Request.Form("SortCatID" & i & "SortFAQID" & j)
					strsql = "UPDATE " & strTablePrefix & "FAQ SET F_FAQ_ORDER=" & SelectNamec & " WHERE F_ID = " & SelectIDc & " AND F_FAQ_CATEGORY = " & SelectID 
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					
					j = j + 1
				Loop
			End If
			i = i + 1
		Loop
	End If
	
	Call OkMessage("F.A.Q. Order Changed!","admin_faq_order.asp","Back to F.A.Q. Order Configuration.")
	Response.Write	"<script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
Else
	Response.Write	"<form action=""admin_faq_order.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewline & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine
	
	strsql = "SELECT FCAT_ID, FCAT_TITLE, FCAT_ORDER FROM " & strTablePrefix & "FAQ_CATEGORY ORDER BY FCAT_ORDER, FCAT_TITLE "
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	If strDBType = "mysql" Then
		strsql2 = "SELECT COUNT(FCAT_ID) AS PAGECOUNT FROM " & strTablePrefix & "FAQ_CATEGORY"
		Set rsCount = my_Conn.Execute(strSql2)
		categorycount = rsCount("PAGECOUNT")
		rsCount.close
		rs.open strSql, my_Conn, adOpenStatic
	Else
		rs.cachesize = 20
		rs.open strSql, my_Conn, adOpenStatic
		If Not (rs.EOF or rs.BOF) Then
			rs.movefirst
			rs.pagesize = 1
			categorycount = cLng(rs.pagecount)
		End If
	End If
	
	Response.Write	"<input name=""NumberCategories"" type=""hidden"" value=""" & categorycount & """>"  & vbNewline & _
					"<table class=""admin"">" & vbNewline & _
					"<tr class=""header"">" & vbNewline & _
					"<td colspan=""2"">Category/F.A.Q.s Order Configuration</td>" & vbNewline & _
					"</tr>" & vbNewline
	
	If rs.EOF Or rs.BOF Then
		Response.Write  "<tr>" & vbNewline & _
						"<td colspan=""2"">No Categories/F.A.Q.s Found</b></td>" & vbNewline & _
						"</tr>" & vbNewline
	Else
		catordercount = 1
		Do Until rs.EOF 
			strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ORDER, F_FAQ_CATEGORY FROM " & strTablePrefix & "FAQ WHERE F_FAQ_CATEGORY=" & rs("FCAT_ID") & " ORDER BY F_FAQ_ORDER ASC, F_FAQ_QUESTION ASC"
			Set rsFAQ = Server.CreateObject("ADODB.Recordset")
			rsFAQ.open strSql, my_Conn, adOpenStatic
			
			If NOT (rsFAQ.EOF or rsFAQ.BOF) Then
				rsFAQ.movefirst
				rsFAQ.pagesize = 1
			End If
			
			If strDBType = "mysql" Then
				strsql2 = "SELECT COUNT(F.FAQ_ID) AS PAGECOUNT FROM " & strTablePrefix & "FAQ F WHERE F.FAQ_CATEGORY = " & rs("FCAT_ID")
				Set rsCount = my_Conn.Execute(strSql2)
				faqcount = rsCount("PAGECOUNT")
				rsCount.close
				set rsCount = nothing
			Else
				FAQcount = cLng(rsFAQ.pagecount)
			End If
			
			Response.Write"<input name=""NumberFAQs" & rs("FCAT_ID") & """ type=""hidden"" value=""" & FAQcount & """> " & vbNewline
			chkDisplayHeader = true
			If rsFAQ.EOF Or rsFAQ.BOF Then
				Response.Write	"<tr class=""section"">" & vbNewLine & _
								"<td>" & ChkString(rs("FCAT_TITLE"),"display") & "</td>" & vbNewline & _
								"<td class=""options"">" & vbNewLine
				
				SelectName = "SortCategory" & catordercount
				SelectID   = "SortCatID" & catordercount
				
				Response.Write	"<input name=""" & SelectID & """ type=""hidden"" value=""" & rs("FCAT_ID") & """>" & vbNewline & _
								"<select name=""" & SelectName & """>" & vbNewline
				
				i = 1
				Do While i <= categorycount
					Response.Write  "<option value=""" & i & """" & chkSelect(i,rs("FCAT_ORDER")) & ">" & i & "</option>" & vbNewline
					i = i + 1
				Loop 
				
				Response.Write	"</select></td>" & vbNewline & _
								"</tr>" & vbNewline & _
								"<tr>" & vbNewline & _
								"<td colspan=""2"">No F.A.Q. Found</td>" & vbNewline & _
								"</tr>" & vbNewline
			Else
				FAQordercount = 1
				Do Until rsFAQ.EOF
					If chkDisplayHeader Then
						Response.Write	"<tr class=""section"">" & vbNewline & _
										"<td>" & ChkString(rs("FCAT_TITLE"),"display") & "</td>" & vbNewline & _
										"<td class=""options"">" & vbNewLine
						
						SelectName = "SortCategory" & catordercount
						SelectID = "SortCatID" & catordercount
						
						Response.Write	"<input name=""" & SelectID & """ type=""hidden"" value=""" & rs("FCAT_ID") & """>" & vbNewline & _
										"<select name=""" & SelectName & """>" & vbNewline
						
						i = 1
						Do While i <= categorycount
							Response.Write  "<option value=""" & i & """" & chkSelect(i,rs("FCAT_ORDER")) & ">" & i & "</option>" & vbNewline
							i = i + 1
						Loop 
						
						Response.Write	"</select>" & vbNewLine & _
										"</td>" & vbNewline & _
										"</tr>" & vbNewline
						chkDisplayHeader = false
					End If
					
					Response.Write  "<tr>" & vbNewline & _
									"<td>" & strType & "&nbsp;" & ChkString(rsFAQ("F_FAQ_QUESTION"),"display") & "</td>" & vbNewline & _
									"<td class=""options"">" & vbNewline
					
					SelectName = "SortCat" & catordercount & "SortFAQ" & FAQordercount
					SelectID   = "SortCatID" & catordercount & "SortFAQID" & FAQordercount
					
					Response.Write	"<input name=""" & SelectID & """ type=""hidden"" value=""" & rsFAQ("F_ID") & """>" & vbNewline & _
									"<select name=""" & SelectName & """>" & vbNewline
					i = 1
					Do While i <= FAQcount
						Response.Write  "<option value=""" & i & """" & chkSelect(i,rsFAQ("F_FAQ_ORDER")) & ">" & i & "</option>" & vbNewline
						i = i + 1
					Loop 
					Response.Write	"</select>" & vbNewLine & _
									"</td>" & vbNewline & _
									"</tr>" & vbNewline
					FAQordercount = FAQordercount + 1
					rsFAQ.MoveNext
				Loop
			End If
			catordercount = catordercount + 1	
			rs.MoveNext
		Loop
		
		rsFAQ.close
		Set rsFAQ = nothing 
		
		Response.Write	"<tr>" & _
						"<td class=""options"" colspan=""2""><button type=""submit"" id=""submit1"" name=""submit1"">Submit Order</button></td>" & vbNewline & _
						"</tr>" & vbNewline
	End If 
	
	Response.Write	"</table>" & vbNewline & _
					"</form>" & vbNewline
	
	rs.close
	Set rs = nothing 
End If
WriteFooterShort
Response.End
%>