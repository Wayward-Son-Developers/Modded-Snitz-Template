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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login_short.asp?target=" & scriptname(ubound(scriptname))
end if
strRqMethod = trim(chkString(Request.QueryString("method"),"SQLString"))
intUsernameID = trim(chkString(Request.QueryString("N_ID"),"SQLString"))

if intUsernameID <> "" then
	if isNumeric(intUsernameID) <> True then intUsernameID = "0"
end if

strPageSize = 10

mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)

Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
				"    <!-----" & vbNewLine & _
				"    function jumpToPage(s) {location.href = s.options[s.selectedIndex].value}" & vbNewLine & _
				"    // -->" & vbNewLine & _
				"</script>" & vbNewLine

Select Case strRqMethod
	Case "Add"
		if Request.Form("Method_Type") = "Write_Configuration" then 
			Err_Msg = ""

			if not IsValidString(trim(Request.Form("strUserName"))) then
				Err_Msg = Err_Msg & "<li>None of the following characters can be used in the username  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
			end if

			txtUserName = chkString(Request.Form("strUserName"),"SQLString")
	
			if txtUserName = " " then 
				Err_Msg = Err_Msg & "<li>You Must Enter a UserName to filter.</li>"
			end if

			if (Instr(txtUserName, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the UserName.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "INSERT INTO " & strFilterTablePrefix & "NAMEFILTER ("
				strSql = strSql & "N_NAME"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & txtUserName & "'"
				strSql = strSql & ")"

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Application.Lock
				Application(strCookieURL & "STRFILTERUSERNAMES") = ""
				Application.UnLock

				Call OkMessage("UserName Added!","admin_config_namefilter.asp","Back To UserName Filter Configuration")
			else
				Call FailMessage(Err_Msg,True)
			end if
		end if
	Case "Delete"
		if Request.Form("Method_Type") = "Delete_UserName" then
			'## Forum_SQL - Delete UserName from NameFilter table
			strSql = "DELETE FROM " & strFilterTablePrefix & "NAMEFILTER "
			strSql = strSql & " WHERE N_ID = " & Request.Form("N_ID")

               		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			Application.Lock
			Application(strCookieURL & "STRFILTERUSERNAMES") = ""
			Application.UnLock

			Call OkMessage("UserName Deleted!","admin_config_namefilter.asp","Back To UserName Filter Configuration")
		else
			Response.Write	"<div class=""warning"" style=""width:50%;"">" & vbNewLine & _
							"<form action=""admin_config_namefilter.asp?method=Delete"" method=""post"" id=""DeleteUserName"" name=""DeleteUserName"">" & vbNewLine & _
							"<input type=""hidden"" name=""Method_Type"" value=""Delete_UserName"">" & vbNewLine & _
							"<input type=""hidden"" name=""N_ID"" value=""" & intUsernameID & """>" & vbNewLine & _
							"<p>Are you sure?</p>" & vbNewLine & _
							"<p><button type=""submit"" id=""submit1"" name=""submit1"">Yes</button>&nbsp;<button onClick=""history.go(-1);"">No</button></p>" & vbNewLine & _
							"</form>" & vbNewLine & _
							"</div>" & vbNewLine
		end if
	Case "Edit"
		if Request.Form("Method_Type") = "Write_Configuration" then
			Err_Msg = ""

			if not IsValidString(trim(Request.Form("strUserName"))) then
				Err_Msg = Err_Msg & "<li>None of the following characters can be used in the username  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
			end if

			txtUserName = chkString(Request.Form("strUserName"),"SQLString")

			if txtUserName = " " then 
				Err_Msg = Err_Msg & "<li>You Must Enter a UserName.</li>"
			end if

			if (Instr(txtUserName, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the UserName.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "UPDATE " & strFilterTablePrefix & "NAMEFILTER "
				strSql = strSql & " SET N_NAME = '" & txtUserName & "'"
				strSql = strSql & " WHERE N_ID = " & Request.Form("N_ID")

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Application.Lock
				Application(strCookieURL & "STRFILTERUSERNAMES") = ""
				Application.UnLock

				Call OkMessage("UserName Filter Updated!","admin_config_namefilter.asp","Back To UserName Filter Configuration")
			else
				Call FailMessage(Err_Msg,True)
			end if
		else
			'## Forum_SQL - Get UserName from DB
			strSql = "SELECT N_ID, N_NAME "
			strSql = strSql & " FROM " & strFilterTablePrefix & "NAMEFILTER "
			strSql = strSql & " WHERE N_ID = " & intUsernameID

			set rs = my_Conn.Execute (strSql)

		        TxtUserName = rs("N_NAME")
		        intN_ID = rs("N_ID")

			rs.close
			set rs = nothing

			Response.Write	"<form action=""admin_config_namefilter.asp?method=Edit"" method=""post"" id=""UpdateUserName"" name=""UpdateUserName"">" & vbNewLine & _
					"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"<input type=""hidden"" name=""N_ID"" value=""" & intN_ID & """>" & vbNewLine & _
					"<table class=""admin"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td>Edit UserName</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>Username</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td><input size=""25"" maxLength=""25"" name=""strUserName"" value=""" & TxtUserName & """ tabindex=""1""></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options""><button type=""submit"" id=""submit1"" name=""submit1"" tabindex=""2"">Update</button></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine & _
					"<p class=""footernav""><a href=""admin_config_namefilter.asp"">Back To UserName Filter Configuration</a></p>" & vbNewLine
		end if
	Case Else
		'## Forum_SQL - Get UserNames from DB
		strSql = "SELECT N_ID, N_NAME " 
		strSql2 = " FROM " & strFilterTablePrefix & "NAMEFILTER "
		strSql3 = " ORDER BY N_NAME ASC "

		if strDBType = "mysql" then 'MySql specific code
			if mypage > 1 then 
				OffSet = cLng((mypage - 1) * strPageSize)
				strSql4 = " LIMIT " & OffSet & ", " & strPageSize & " "
			end if

			'## Forum_SQL - Get the total pagecount 
			strSql1 = "SELECT COUNT(N_ID) AS PAGECOUNT "

			set rsCount = my_Conn.Execute(strSql1 & strSql2)
			iPageTotal = rsCount(0).value
			rsCount.close
			set rsCount = nothing

			If iPageTotal > 0 then
				maxpages = (iPageTotal \ strPageSize )
				if iPageTotal mod strPageSize <> 0 then
					maxpages = maxpages + 1
				end if
				if iPageTotal < (strPageSize + 1) then
					intGetRows = iPageTotal
				elseif (mypage * strPageSize) > iPageTotal then
					intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
				else
					intGetRows = strPageSize
				end if
			else
				iPageTotal = 0
				maxpages = 0
			end if 

			if iPageTotal > 0 then
				set rs = Server.CreateObject("ADODB.Recordset")
				rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
					arrUsernameData = rs.GetRows(intGetRows)
					iUsernameCount = UBound(arrUsernameData, 2)
				rs.close
				set rs = nothing
			else
				iUsernameCount = ""
			end if
 
		else 'end MySql specific code

			set rs = Server.CreateObject("ADODB.Recordset")
			rs.cachesize = strPageSize
			rs.open strSql & strSql2 & strSql3, my_Conn, adOpenStatic
				If not (rs.EOF or rs.BOF) then
					rs.movefirst
					rs.pagesize = strPageSize
					rs.absolutepage = mypage '**
					maxpages = cLng(rs.pagecount)
					arrUsernameData = rs.GetRows(strPageSize)
					iUsernameCount = UBound(arrUsernameData, 2)
				else
					iUsernameCount = ""
				end if
			rs.Close
			set rs = nothing
		end if

		Response.Write	"<form action=""admin_config_namefilter.asp?method=Add"" method=""post"" id=""Add"" name=""Add"">" & vbNewLine & _
						"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
						"<table class=""admin"">" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td colspan=""2"">UserName Filter Configuration</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""section"">" & vbNewLine & _
						"<td>UserName</td>" & vbNewLine
		if maxpages > 1 then
			Call DropDownPaging()
		else
			Response.Write	"<td>&nbsp;</td>" & vbNewLine
		end if
		Response.Write	"</tr>" & vbNewLine

		if iUsernameCount = "" then  '## No Badwords found in DB
			Response.Write	"<tr>" & vbNewLine & _
							"<td colspan=""2"">No UserNames Found</td>" & vbNewLine & _
							"</tr>" & vbNewLine
		else
			nN_ID = 0
			nN_NAME = 1

			rec = 1
			intI = 0

			for iUsername = 0 to iUsernameCount
				if (rec = strPageSize + 1) then exit for

				Username_ID = arrUsernameData(nN_ID, iUsername)
				Username_Name = arrUsernameData(nN_NAME, iUsername)

				if intI = 1 then 
					CColor = strAltForumCellColor
				else
					CColor = strForumCellColor
				end if

				Response.Write	"<tr>" & vbNewLine & _
								"<td>" & Username_Name & "</td>" & vbNewLine & _
								"<td class=""options""><a href=""admin_config_namefilter.asp?method=Edit&N_ID=" & Username_ID & """>" & getCurrentIcon(strIconPencil,"Edit UserName","hspace=""0""") & "</a>&nbsp;<a href=""admin_config_namefilter.asp?method=Delete&N_ID=" & Username_ID & """>" & getCurrentIcon(strIconTrashcan,"Delete UserName","hspace=""0""") & "</a></td>" & vbNewLine & _
								"</tr>" & vbNewLine
				rec = rec + 1
				intI = intI + 1
				if intI = 2 then
					intI = 0
				end if
			next
		end if

		Response.Write	"<tr>" & vbNewLine & _
				"<td><input size=""25"" maxLength=""25"" name=""strUserName"" value=""" & TxtUserName & """ tabindex=""1""></td>" & vbNewLine & _
				"<td class=""options""><button type=""submit"" tabindex=""2"">Add</button></td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"</form>" & vbNewLine
End Select
WriteFooterShort
Response.End

sub DropDownPaging()
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		Response.Write	"<td class=""options"">" & vbNewLine & _
						"<b>Page</b>&nbsp;<select style=""font-size:9px"" name=""whichpage"" size=""1"" onchange=""jumpToPage(this)"">" & vbNewLine
		for counter = 1 to maxpages
			ref = "admin_config_namefilter.asp?whichpage=" & counter 
			if counter <> cLng(pge) then
				Response.Write	"<option value=""" & ref & """>" & counter & "</option>" & vbNewLine
			else
				Response.Write	"<option value=""" & ref & """ selected>" & counter & "</option>" & vbNewLine
			end if
		next
		Response.Write	"</select>&nbsp;<b>of " & maxpages & "</b></td>" & vbNewLine
	end if
end sub 

Function IsValidString(sValidate)
	Dim sInvalidChars
	Dim bTemp
	Dim i 
	' Disallowed characters
	sInvalidChars = "!#$%^&*()=+{}[]|\;:/?>,<'"
	for i = 1 To Len(sInvalidChars)
		if InStr(sValidate, Mid(sInvalidChars, i, 1)) > 0 then bTemp = True
		if bTemp then Exit For
	next
	for i = 1 to Len(sValidate)
		if Asc(Mid(sValidate, i, 1)) = 160 then bTemp = True
		if bTemp then Exit For
	next

	' extra checks
	' no two consecutive dots or spaces
	if not bTemp then
		bTemp = InStr(sValidate, "..") > 0
	end if
	if not bTemp then
		bTemp = InStr(sValidate, "  ") > 0
	end if
	if not bTemp then
		bTemp = (len(sValidate) <> len(Trim(sValidate)))
	end if 'Addition for leading and trailing spaces

	' if any of the above are true, invalid string
	IsValidString = Not bTemp
End Function
%>
