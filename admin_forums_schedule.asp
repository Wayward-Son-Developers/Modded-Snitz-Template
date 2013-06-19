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
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Forum Deletion/Archival<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table><br /><center>" & vbNewLine

strWhatToDo = request("action")
if strWhatToDo = "" then
	strWhatToDo = "default"
end if

Select Case strWhatToDo
	Case "updateArchive"
		if Request("id") = "" or IsNull(Request("id")) then
			Response.write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>There has been a problem!</font></p>" & vbNewLine & _
					"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>No Forums Selected!</font></p>" & vbNewLine & _
					"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go back to correct the problem.</a></font></p>" & vbNewLine
			WriteFooter
			Response.End
		end if
		Response.Write	"      <table border=""0"" width=""75%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Administrative Forum Archive Schedule</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""top""><br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><ul>" & vbNewLine
		reqID = split(Request.Form("id"), ",")
		for i = 0 to ubound(reqID)
			tmpStr = "archSched" & trim(reqID(i))
			if tmpStr = "" then tmpStr = NULL
			strSQL = "UPDATE " & strTablePrefix & "FORUM SET F_ARCHIVE_SCHED = " & cLng("0" & Request.Form(tmpStr))
			strSQL = strSQL & " WHERE FORUM_ID = " & cLng("0" & trim(reqID(i)))
			my_conn.execute(strSQL),,adCmdText + adExecuteNoRecords
			Response.Write	"                <li>Archive Schedule for <b>" & GetForumName(reqID(i)) & "</b> updated to " & Request.Form(tmpStr) & " days.</li>" & vbNewLine
		next
		Response.Write	"                </ul></font><br /></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <br />" & vbNewLine & _
				"      <div align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
				"      <a href=""admin_forums.asp"">Back to Forums Administration</a></font></div><br />" & vbNewLine
	Case "default" '################ ARCHIVE
		Response.Write	"      <table border=""0"" width=""75%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"" >" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Administrative Forum Archive Functions</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><b>Archive Reminder:</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""top"">" & vbNewLine
		strForumIDN = request("id")
		if strForumIDN = "" then
			strsql = "SELECT CAT_ID, FORUM_ID, F_L_ARCHIVE,F_ARCHIVE_SCHED, F_SUBJECT FROM " & strTablePrefix & "FORUM ORDER BY CAT_ID, F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)    
			thisCat = 0
			if drs.eof then
				Response.Write	"                  No Forums Found!" & vbNewLine
			else
				Response.Write	"                  <form name=""arcTopic"" action=""admin_forums_schedule.asp"" method=""post"">" & vbNewLine & _
						"                  <input type=""hidden"" name=""action"" value=""updateArchive"">" & vbNewLine & _
						"                  <table width=""100%"" cellpadding=""1"" cellspacing=""3"">" & vbNewLine
				do until drs.eof
					if (IsNull(drs("F_L_ARCHIVE"))) or (drs("F_L_ARCHIVE") = "") then 
						archive_date = "Not archived" 
					else 
						archive_date = StrToDate(drs("F_L_ARCHIVE"))
					end if

					Response.Write	"                    <tr>" & vbNewLine & _
							"                      <td bgcolor=""" & strForumCellColor & """ valign=""bottom""><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & drs("F_SUBJECT") & "</font></td>" & vbNewLine & _
							"                      <td bgcolor=""" & strForumCellColor & """ valign=""bottom"" align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """> archive schedule: " & "<input type=""text"" name=""archSched" & Trim(drs("FORUM_ID")) & """ size=""3"" value=""" & drs("F_ARCHIVE_SCHED") & """ maxlength=""3""> days" & "</font></td>" & vbNewLine & _
							"                    </tr>" & vbNewLine
					thisCat = drs("CAT_ID")
					drs.movenext
				loop
				Response.Write	"                    <tr>" & vbNewLine & _
						"                      <td bgcolor=""" & strForumCellColor & """ colspan=""2""><center><input type=""submit"" name=""submit1"" value=""Update Schedule""></center></td>" & vbNewLine & _
						"                    </tr>" & vbNewLine & _
						"                  </table>" & vbNewLine & _
						"                  </form>" & vbNewLine
			end if
			set drs = nothing
			Response.Write	"                </td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"            </table>" & vbNewLine & _
					"          </td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      <br /><a href=""admin_forums.asp""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Back to Forums Administration</font></a><br /><br />" & vbNewLine
		end if
end Select
WriteFooter
Response.End

Function GetForumName(fID)
	'## Forum_SQL
	strSql = "SELECT F.F_SUBJECT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM F"
	strSql = strSql & " WHERE F.FORUM_ID = " & fID

	set rsGetForumName = my_Conn.Execute(strSql)

	if rsGetForumName.bof or rsGetForumName.eof then
		GetForumName = ""
	else
		GetForumName = rsGetForumName("F_SUBJECT")
	end if

	set rsGetForumName = nothing
end Function
%>
