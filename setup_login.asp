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

blnSetup = Request.Form("setup")
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_func_common.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<%
Response.Write	"<html>" & vbNewLine & _
		vbNewLine & _
		"<head>" & vbNewLine & _
		"<title>Forum-Setup Page</title>" & vbNewLine

'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser"">" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

Response.Write	"<style><!--" & vbNewLine & _
		"a:link    {color:darkblue;text-decoration:underline}" & vbNewLine & _
		"a:visited {color:blue;text-decoration:underline}" & vbNewLine & _
		"a:hover   {color:red;text-decoration:underline}" & vbNewLine & _
		"--></style>" & vbNewLine & _
		"</head>" & vbNewLine & _
		vbNewLine & _
		"<body bgColor=""white"" text=""midnightblue"" link=""darkblue"" aLink=""red"" vLink=""red"" onLoad=""window.focus()"">" & vbNewLine

set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

Name = trim(chkString(Request.Form("Name"),"SQLString"))
Password = trim(chkString(Request.Form("Password"),"SQLString"))
ReturnTo = Request.Form("ReturnTo")

RequestMethod = Request.ServerVariables("Request_method")

if RequestMethod = "POST" Then
	'## Forum_SQL
	strSql = "SELECT COUNT(*) AS ApprovalCode "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & Name & "' AND ("
	strSql = strSql & "       M_PASSWORD = '" & Password & "' OR M_PASSWORD = '" & sha256("" & Password) & "') AND "
	strSql = strSql & "       M_LEVEL = 3"
	
	set dbRs = my_Conn.Execute(strSql)

	if dbRS.Fields("ApprovalCode") = "1"  and ChkQuoteOk(Name) and ChkQuoteOk(Password) then
		Response.Write	"<p>&nbsp;</p>" & vbNewLine & _
				"<p align=""center""><font face=""Verdana, Arial, Helvetica"" size=""4"">Login was successful!</font></p>" & vbNewLine
		Session(strCookieURL & "Approval") = strAdminCode
		Response.Write	"<p>&nbsp;</p>" & vbNewLine & _
				"<p align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp?" & Server.URLEncode(ReturnTo) & """ target=""_top"">Click here to Continue.</a></font></p>" & vbNewLine & _
				"<meta http-equiv=""Refresh"" content=""2; URL=setup.asp?" & Server.URLEncode(ReturnTo) & """>" & vbNewLine
		Response.End
	else
		Response.Write	"<div align=""center""><center>" & vbNewLine & _
				"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">There has been a problem !</font></p>" & vbNewLine & _
				"</center></div>" & vbNewLine & _
				"<form action=""setup_login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"<input type=""hidden"" name=""setup"" value=""Y"">" & vbNewLine & _
				"<input type=""hidden"" name=""ReturnTo"" value=""" & Request.Form("ReturnTo") & """>" & vbNewLine & _
				"<table width=""50%"" height=""50%"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""5"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""center""><p align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>You are not allowed access.</b></font></p></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""left""><p><font face=""Verdana, Arial, Helvetica"" size=""2"">If you think you have reached this message in error, please try again.</font></p></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine & _
				"      <table border=""0"" cellspacing=""2"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""center"" colspan=""2"" bgColor=""#9FAFDF""><b><font face=""Verdana, Arial, Helvetica"" size=""2"">Admin Login</font></b></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""right"" nowrap><b><font face=""Verdana, Arial, Helvetica"" size=""2"">UserName:</font></b></td>" & vbNewLine & _
				"          <td><input type=""text"" name=""Name""></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""right"" nowrap><b><font face=""Verdana, Arial, Helvetica"" size=""2"">Password:</font></b></td>" & vbNewLine & _
				"          <td><input type=""Password"" name=""Password""></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td colspan=""2"" align=""right""><input type=""submit"" value=""Login"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"</form>" & vbNewLine
	end if
	set dbRS = nothing
else
	Response.Redirect("default.asp")
end if

my_Conn.close
set my_Conn = nothing

Response.Write	"</body>" & vbNewLine & _
		vbNewLine & _
		"</html>" & vbNewLine
%>
