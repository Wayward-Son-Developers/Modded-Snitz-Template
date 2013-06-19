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

'#################################################################################
    strNewVersion = "Snitz Forums 2000 Version 3.4.07"
'#################################################################################
Dim NewConfig
ResponseCode = Request.QueryString("RC")
%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<%
Dim strCurrentDateTime
Dim strlhDateTime
strCurrentDateTime = DateToStr(Now())
strlhDateTime = DateToStr(dateadd("n", -5, Now()))

if ResponseCode <> "" then 'No parameter
	blnSetup = "Y"
else
	strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))
	Application.Lock
	Application(strCookieURL & "ConfigLoaded")= ""
	Application.UnLock
end if
if blnSetup <> "Y" then NewConfig = 1
%>
<!-- #INCLUDE FILE="config.asp" -->
<%

Response.Buffer = True

Response.Write	"<html>" & vbNewLine & _
		vbNewLine & _
		"<head>" & vbNewLine & _
		"<title>Forum-Setup Page</title>" & vbNewLine

'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

Response.Write	"<style><!--" & vbNewLine & _
		"a:link    {color:darkblue;text-decoration:underline}" & vbNewLine & _
		"a:visited {color:blue;text-decoration:underline}" & vbNewLine & _
		"a:hover   {color:red;text-decoration:underline}" & vbNewLine & _
		"--></style>" & vbNewLine & _
		"</head>" & vbNewLine & _
		vbNewLine & _
		"<body bgColor=""white"" text=""midnightblue"" link=""darkblue"" aLink=""red"" vLink=""red"" onLoad=""window.focus()"">" & vbNewLine

If strDBType = "" then 
	Response.Write	"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""center""><p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
			"<b>Database Setup....</b><br /><br />" & _
			"Your <b>strDBType</b> is not set, please edit your <b>config.asp</b><br />to reflect your database type." & _
			"</font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
			"<a href=""setup.asp"" target=""_top"">Click here to retry.</a></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</body>" & vbNewLine & _
			"</html>" & vbNewLine
	Response.End
end if

if ResponseCode = "" then 'No parameter

'	Check to see if all the fields are in the database

	on error resume next

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDesc = my_conn.Errors(counter).Description
		if ConnErrorNumber <> 0 then 
			my_Conn.Errors.Clear
			Err.Clear
			Response.Redirect "setup.asp?RC=1&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc)
		end if
	next

	my_Conn.Errors.Clear
	Err.Clear

	strSql = "SELECT CAT_ID, FORUM_ID, F_STATUS, F_MAIL, F_SUBJECT, F_URL, F_DESCRIPTION, F_TOPICS, F_COUNT, F_LAST_POST, "
	strSql = strSql & "F_PASSWORD_NEW, F_PRIVATEFORUMS, F_TYPE, F_IP, F_LAST_POST_AUTHOR, F_A_TOPICS, F_A_COUNT, "
	strSQL = strSQL & "F_MODERATION, F_SUBSCRIPTION, F_ORDER, F_L_ARCHIVE, F_ARCHIVE_SCHED, F_L_DELETE, F_DELETE_SCHED"
	strSql = strSql &  " FROM " & strTablePrefix & "FORUM"

	my_Conn.Execute strSql

	Call CheckSqlError()

	my_Conn.Errors.Clear
	Err.Clear

	strSql = "SELECT CAT_ID, FORUM_ID, TOPIC_ID, T_STATUS, T_SUBJECT, T_MESSAGE, T_AUTHOR, T_REPLIES, "
	strSql = strSql & " T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_ARCHIVE_FLAG, T_LAST_POST_AUTHOR "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS"

	my_Conn.Execute strSql

	Call CheckSqlError()

	my_Conn.Errors.Clear
	Err.Clear

	strSql = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql = strSql & "M_SIG, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_ICQ, M_MSN, M_YAHOO, M_POSTS, M_DATE, M_LASTHEREDATE, "
	strSql = strSql & "M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, M_HIDE_EMAIL, M_RECEIVE_EMAIL, M_LAST_IP, M_IP, "
	strSql = strSql & "M_FIRSTNAME, M_LASTNAME, M_OCCUPATION, M_SEX, M_AGE, M_HOBBIES, M_LNEWS, M_QUOTE, M_BIO, "
	strSql = strSql & "M_MARSTATUS, M_LINK1, M_LINK2, M_CITY, M_STATE, M_PHOTO_URL, M_KEY, M_NEWEMAIL, M_PWKEY, M_SHA256 "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"

	my_Conn.Execute strSql

	Call CheckSqlError()

	my_Conn.Errors.Clear
	Err.Clear

	strSql = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql = strSql & "M_SIG, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_ICQ, M_MSN, M_YAHOO, M_POSTS, M_DATE, M_LASTHEREDATE, "
	strSql = strSql & "M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, M_HIDE_EMAIL, M_RECEIVE_EMAIL, M_LAST_IP, M_IP, "
	strSql = strSql & "M_FIRSTNAME, M_LASTNAME, M_OCCUPATION, M_SEX, M_AGE, M_HOBBIES, M_LNEWS, M_QUOTE, M_BIO, "
	strSql = strSql & "M_MARSTATUS, M_LINK1, M_LINK2, M_CITY, M_STATE, M_PHOTO_URL, M_KEY, M_NEWEMAIL, M_PWKEY, M_SHA256 "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"

	my_Conn.Execute strSql

	Call CheckSqlError()

	my_Conn.Errors.Clear
	Err.Clear

	on error goto 0

	if strVersion <> strNewVersion then
		Response.Redirect "setup.asp?RC=3&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & Server.UrlEncode("Different or New Version-ID detected")
	end if

	'## This part of the code is only reached if all is ok !!

	Response.Write	"<div align=""center""><center>" & vbNewLine & _
			"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">Forum setup has been completed.</font></p>" & vbNewLine & _
			"</center></div>" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    <b>Congratulations!!</b><br />" & vbNewLine & _
			"    The forum setup has been completed successfully.<br />" & vbNewLine & _
			"    You can now start using Snitz Forums 2000.</font></p>" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    If you have questions or remarks you can visit us at: <a href=""http://forum.snitz.com"">http://forum.snitz.com</a></font></p>" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    You can also post the address of your forum there<br />" & vbNewLine & _
			"    so others can come and visit you." & vbNewLine & _
			"    </font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center"">" & vbNewLine & _
			"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    <a href=""default.asp"" target=""_top"">Click here to go to the forum.</a>" & vbNewLine & _
			"    </font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center"">" & vbNewLine & _
			"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    <a href=""setup.asp?RC=3"" target=""_top"">Upgrade the database.</a><br /><small>(shouldn't be needed for this database!)</small>" & vbNewLine & _
			"    </font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine
	if strDBType <> "access" then
		Response.Write	"  <tr>" & vbNewLine & _
				"    <td align=""center"">" & vbNewLine & _
				"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    <a href=""setup.asp?RC=5"" target=""_top"">Create the database tables.</a><br /><small>(shouldn't be needed for this database!)</small>" & vbNewLine & _
				"    </font></td>" & vbNewLine & _
				"  </tr>" & vbNewLine
	end if
	Response.Write	"</table>" & vbNewLine

elseif ResponseCode = 1 then '## cannot open database

	ErrorCode = Request.QueryString("EC")
	ErrorDesc = Request.QueryString("ED")
	CustomCode = Request.QueryString("CC")

	Response.Write	"<div align=""center""><center>" & vbNewLine & _
			"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">There has been an error !!</font></p>" & vbNewLine & _
			"</center></div>" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine
	if CustomCode = 1 then
		Response.Write	"    The database could not be opened !!<br />" & vbNewLine & _
				"    Check your config.asp file and set the <br /><b>strConnString</b> so it points to the database.<br />" & vbNewLine & _
				"    Also check if <b>strDBType</b> is set to the right databasetype.<br />" & vbNewLine & _
				"    <br />" & vbNewLine
	elseif CustomCode = 2 then
		Response.Write	"    Couldn't read from one or more tables in the database.<br /> Make sure none of the tables are exclusively locked by another user.<br /><br />" & vbNewLine
	elseif CustomCode = 3 then
		Response.Write	"    Couldn't open the database.<br /> Make sure you supplied a correct username and password.<br /><br />" & vbNewLine
	else
		Response.Write	"    The database could not be opened !!<br />" & vbNewLine & _
				"    <br />" & vbNewLine
	end if
	if ErrorCode <> "" and ErrorCode < "0" then
		Response.Write("    <p>Code :  " & Hex(ErrorCode) & "</p>" & vbNewLine)
		if ErrorDesc <> "" then
			Response.Write("	<p><b>Error Description</b> : <br />" & ErrorDesc & "</p>" & vbNewLine)
		end if
	elseif ErrorCode <> "" then
		Response.Write("    <p>Code :  " & ErrorCode & "</p>" & vbNewLine)
		if ErrorDesc <> "" then
			Response.Write("	<p><b>Error Description</b> : <br />" & ErrorDesc & "</p>" & vbNewLine)
		end if
	end if
	Response.Write	"    </font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center"">" & vbNewLine & _
			"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine
	if Request.QueryString("RET") <> "" then
		Response.Write	"    <a href=""" & Request.QueryString("RET") & """ target=""_top"">Click here to return to the previous screen.</a>" & vbNewLine
	else
		Response.Write	"    <a href=""setup.asp"" target=""_top"">Click here to retry.</a>" & vbNewLine
	end if
	Response.Write	"    </font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine

elseif ResponseCode = 2 then '## cannot find all the fields in the database

	strSender = Request.QueryString("MAIL")
	strVersion = Request.QueryString("VER")
	ErrorCode = Request.QueryString("EC")
	CustomCode = Request.QueryString("CC")

	if ErrorCode = "-2147467259" then
		if strVersion <> "" then
			Response.Redirect "setup.asp?RC=3&VER=" & strVersion
			Response.End
		else
			Response.Redirect "setup.asp?RC=5&EC=" & ErrorCode
			Response.End
		end if
	elseif ErrorCode = "-2147217865" then
		if strVersion <> "" then
			Response.Redirect "setup.asp?RC=3&VER=" & strVersion
			Response.End
		end if
		Response.Write	"<div align=""center""><center>" & vbNewLine & _
				"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">The database needs to be installed !!</font></p>" & vbNewLine & _
				"</center></div>" & vbNewLine & _
				"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    You need to create all the tables in the database before you can start using the forum.<br />" & vbNewLine & _
				"    </font></p></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td align=""center"">" & vbNewLine & _
				"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    <a href=""setup.asp?RC=5&strDBType=" & strDBType & """ target=""_top"">Click here to create the tables in the database.</a><br /><br />" & vbNewLine & _
				"    <a href=""setup.asp"" target=""_top"">Click here to retry.</a>" & vbNewLine & _
				"    </font></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine
	else
		Response.Write	"<div align=""center""><center>" & vbNewLine & _
				"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">The database needs to be upgraded !!</font></p>" & vbNewLine & _
				"</center></div>" & vbNewLine & _
				"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    The database you are using needs to be upgraded !!<br />" & vbNewLine
		if MAIL <> "" then
			Response.Write	"    If you are not an Administrator at this forum<br /> please report this error here: <a href=""mailto:" & strSender & """>" & strSender & "</a>.<br /><br />" & vbNewLine
		end if
		if ErrorCode <> "" and ErrorCode < "0" then
			Response.Write("    <p>Code :  " & Hex(ErrorCode) & "</p>" & vbNewLine)
		elseif ErrorCode <> "" then
			Response.Write("    <p>Code :  " & ErrorCode & "</p>" & vbNewLine)
		end if
		Response.Write	"    </font></p></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td align=""center"">" & vbNewLine & _
				"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    <a href=""setup.asp?RC=3&MAIL=" & Server.URLEncode(strSender) & """ target=""_top"">Click here to upgrade the database.</a><br /><br />" & vbNewLine & _
				"    <a href=""default.asp"" target=""_top"">Click here to retry.</a>" & vbNewLine & _
				"    </font></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine
	end if

elseif ResponseCode = 3 then '## upgrade database

	if strVersion = "" then
		strVersion = Request.QueryString("VER")
	end if
	if Session(strCookieURL & "Approval") = strAdminCode then

		'## logon was ok proceed with upgrade
		Response.Write	"<div align=""center""><center>" & vbNewLine
		if strDBType = "sqlserver" then
			Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">Select the SQL-Server upgrade options.</font></p>" & vbNewLine
		elseif strDBType = "mysql" then
			Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">MySql database upgrade.</font></p>" & vbNewLine
		else
			Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">Access 97/2000/2002 database upgrade</font></p>" & vbNewLine
		end if
		Response.Write	"</center></div>" & vbNewLine & _
				"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"""
		if strDBType = "sqlserver" then Response.Write(" align=""left""") else Response.Write(" align=""center""")
		Response.Write	">" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    Select the version you want to upgrade from:" & vbNewLine & _
				"    </font></p>" & vbNewLine & _
				"    <form action=""setup.asp?RC=4&strDBType=" & strDBType & """ method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"    <p><select size=""1"" name=""OldVersion"">" & vbNewLine & _
				"    <option value=""12""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4.06") & ">Snitz Forums 2000 Version 3.4.06</option>" & vbNewLine & _
				"    <option value=""11""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4.05") & ">Snitz Forums 2000 Version 3.4.05</option>" & vbNewLine & _
				"    <option value=""11""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4.04") & ">Snitz Forums 2000 Version 3.4.04</option>" & vbNewLine & _
				"    <option value=""10""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4.03") & ">Snitz Forums 2000 Version 3.4.03</option>" & vbNewLine & _
				"    <option value=""10""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4.02") & ">Snitz Forums 2000 Version 3.4.02</option>" & vbNewLine & _
				"    <option value=""10""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4.01") & ">Snitz Forums 2000 Version 3.4.01</option>" & vbNewLine & _
				"    <option value=""9""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.4") & ">Snitz Forums 2000 Version 3.4</option>" & vbNewLine & _
				"    <option value=""8""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.3.05") & ">Snitz Forums 2000 Version 3.3.05</option>" & vbNewLine & _
				"    <option value=""7""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.3.04") & ">Snitz Forums 2000 Version 3.3.04</option>" & vbNewLine & _
				"    <option value=""7""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.3.03") & ">Snitz Forums 2000 Version 3.3.03</option>" & vbNewLine & _
				"    <option value=""6""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.3.02") & ">Snitz Forums 2000 Version 3.3.02</option>" & vbNewLine & _
				"    <option value=""6""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.3.01") & ">Snitz Forums 2000 Version 3.3.01</option>" & vbNewLine & _
				"    <option value=""6""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.3 Final") & ">Snitz Forums 2000 Version 3.3 Final</option>" & vbNewLine & _
				"    <option value=""5""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.1 SR4") & ">Snitz Forums 2000 V3.1 Service Release 4</option>" & vbNewLine & _
				"    <option value=""5""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.1 SR3 Final") & ">Snitz Forums 2000 V3.1 Service Release 3 Final</option>" & vbNewLine & _
				"    <option value=""5""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.1 SR3b4") & ">Snitz Forums 2000 V3.1 Service Release 3 Beta 4</option>" & vbNewLine & _
				"    <option value=""4"">Snitz Forums 2000 V3.1 Service Release 3 Beta 2</option>" & vbNewLine & _
				"    <option value=""3"">Snitz Forums 2000 V3.1 Service Release 3 Beta 1</option>" & vbNewLine & _
				"    <option value=""3"">Snitz Forums 2000 V3.1 Service Release 2</option>" & vbNewLine & _
				"    <option value=""3""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.1 Service Release 1 ") & ">Snitz Forums 2000 V3.1 Service Release 1</option>" & vbNewLine & _
				"    <option value=""3""" & CheckSelected(strVersion,"Snitz Forums 2000 Version 3.1 final ") & ">Snitz Forums 2000 V3.1 Final</option>" & vbNewLine & _
				"    <option value=""3"">Snitz Forums 2000 V3.1 Beta 5</option>" & vbNewLine & _
				"    <option value=""3"">Snitz Forums 2000 V3.1 Beta 4</option>" & vbNewLine & _
				"    <option value=""2"">Snitz Forums 2000 V3.1 Beta 3</option>" & vbNewLine & _
				"    <option value=""1"">Snitz Forums 2000 V3.1 Beta 2</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3.0 Service Release 2</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3.0 Service Release 1</option>" & vbNewLine & _
				"    <option value=""0""" & CheckSelected(strVersion,"Snitz Forums 2000 v3.0") & ">Snitz Forums 2000 Version 3.0 Final</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3 RC5</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3 RC4</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3 RC3</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3 RC2</option>" & vbNewLine & _
				"    <option value=""0"">Snitz Forums 2000 V3 RC1</option>" & vbNewLine & _
				"    </select></p>" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    Current database versionstring =<br /><b> " & strVersion & "</b>" & vbNewLine & _
				"    </font></p>" & vbNewLine
		if strDBType = "sqlserver" then
			Response.Write	"    <p>Select the SQL-server version you are using:</p>" & vbNewLine & _
					"    <p><input type=""radio"" class=""radio"" name=""SQL_Server"" value=""SQL6"">SQL-Server 6.5<br />" & vbNewLine & _
					"    <input type=""radio"" class=""radio"" checked name=""SQL_Server"" value=""SQL7"">SQL-Server 7/2000</p>" & vbNewLine
		end if
		if strDBType <> "access" then
			Response.Write	"    <p>To upgrade the database you need to provide a username" & vbNewLine & _
					"    and password of a user that has table creation/modification rights at the database you use.<br />" & vbNewLine & _
					"    This might not be the same user as you use in your connectionstring !<br />" & vbNewLine & _
					"    <br />" & vbNewLine & _
					"    Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type=""text"" name=""DBUserName"" size=""20""><br />" & vbNewLine & _
					"    Password: <input type=""password"" name=""DBPassword"" size=""20""></p>" & vbNewLine
		end if
		Response.Write	"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td align=""center""><input type=""submit"" value=""Continue"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</form>" & vbNewLine & _
				"</table>" & vbNewLine
	else
		strSender = Request.QueryString("MAIL")

		Response.Write	"<div align=""center""><center>" & vbNewLine & _
				"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">You need to logon first.</font></p>" & vbNewLine & _
				"</center></div>" & vbNewLine & _
				"<form action=""setup_login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"<input type=""hidden"" name=""setup"" value=""Y"">" & vbNewLine & _
				"<input type=""hidden"" name=""ReturnTo"" value=""RC=3&VER=" & strVersion & """>" & vbNewLine & _
				"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""left"">" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
				"    To upgrade the database you need to be logged on as a forum administrator.<br />" & vbNewLine
		if strSender <> "" then
			Response.Write	"    If you are not the Administrator of this forum<br /> please report this error here: <a href=""mailto:" & strSender & """>" & strSender & "</a>.<br /><br />" & vbNewLine
		end if
		Response.Write	"    </font></p></td>" & vbNewLine & _
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
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"</form>" & vbNewLine & _
				"</font>" & vbNewLine
	end if
	
elseif ResponseCode = 4 then '## start upgrading database

	if Session(strCookieURL & "Approval") = strAdminCode Then

		'## logon was ok proceed with upgrade
		Response.Write	"<div align=""center""><center>" & vbNewLine & _
				"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">Please Wait until the upgrade has been completed !</font></p>" & vbNewLine

		strSQL_Server = Request.Form("Sql_Server")

		if strDBType = "access" or not Instr(strConnString,"uid=") > 0 then
			strUpgradeString = strConnString
		else
			strUpgradeString = CreateConnectionString(strConnString, Request.Form("DBUserName"), Request.Form("DBPassword"))
		end if

		on error resume next

		set my_Conn = Server.CreateObject("ADODB.Connection")
		my_Conn.Open strUpgradeString

		for counter = 0 to my_Conn.Errors.Count -1
			ConnErrorNumber = Err.Number
			ConnErrorDesc = my_conn.Errors(counter).Description
			if ConnErrorNumber <> 0 then 
				my_Conn.Errors.Clear
				Err.Clear 
				Response.Redirect "setup.asp?RC=1&CC=3&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc) & "&RET=" & Server.URLEncode("setup.asp?RC=3")
			end if
		next

		on error goto 0

		dim intCriticalErrors, intWarnings, Prefix, FieldName, TableName, DataType
		intCriticalErrors = 0
		intWarnings = 0

		Prefix	  = 1
		FieldName = 2
		TableName = 3
		DataType_Access  = 4
		DataType_SQL6 = 5
		DataType_SQL7 = 6
		DataType_MySQL = 7 
		ConstraintAccess = 8
		ConstraintSQL6 = 9 
		ConstraintSQL7 = 10
		ConstraintMySQL = 11
		Access = 1
		SQL6 = 2
		SQL7 = 3
		MySql = 4

		if not(IsNull(Request.Form("OldVersion"))) then
			OldVersion = Request.Form("OldVersion")
		else
			OldVersion = Request.QueryString("OldVersion")
		end if

		if OldVersion = 0 then

			Dim NewColumns(8,11)

			NewColumns(0, Prefix)	 = strTablePrefix
			NewColumns(0, FieldName) = "C_STRSHOWSTATISTICS"
			NewColumns(0, TableName) = "CONFIG"
			NewColumns(0, DataType_Access) = "SMALLINT"
			NewColumns(0, DataType_SQL6) = "SMALLINT"
			NewColumns(0, DataType_SQL7) = "SMALLINT"
			NewColumns(0, DataType_MySQL) = "SMALLINT"
			NewColumns(0, ConstraintAccess)  = "NULL"
			NewColumns(0, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0362 DEFAULT 1"
			NewColumns(0, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0362 DEFAULT 1"
			NewColumns(0, ConstraintMySQL)  = "DEFAULT 1 NULL"

			NewColumns(1, Prefix)	 = strTablePrefix
			NewColumns(1, FieldName) = "C_STRSHOWIMAGEPOWEREDBY"
			NewColumns(1, TableName) = "CONFIG"
			NewColumns(1, DataType_Access) = "SMALLINT"
			NewColumns(1, DataType_SQL6) = "SMALLINT"
			NewColumns(1, DataType_SQL7) = "SMALLINT"
			NewColumns(1, DataType_MySQL) = "SMALLINT"
			NewColumns(1, ConstraintAccess)  = "NULL"
			NewColumns(1, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0363 DEFAULT 1"
			NewColumns(1, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0363 DEFAULT 1"
			NewColumns(1, ConstraintMySQL)  = "DEFAULT 1 NULL"

			NewColumns(2, Prefix)	 = strTablePrefix
			NewColumns(2, FieldName) = "C_STRLOGONFORMAIL"
			NewColumns(2, TableName) = "CONFIG"
			NewColumns(2, DataType_Access) = "SMALLINT"
			NewColumns(2, DataType_SQL6) = "SMALLINT"
			NewColumns(2, DataType_SQL7) = "SMALLINT"
			NewColumns(2, DataType_MySQL) = "SMALLINT"
			NewColumns(2, ConstraintAccess)  = "NULL"
			NewColumns(2, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0364 DEFAULT 1"
			NewColumns(2, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0364 DEFAULT 1"
			NewColumns(2, ConstraintMySQL)  = "DEFAULT 1 NULL"

			NewColumns(3, Prefix)	 = strTablePrefix
			NewColumns(3, FieldName) = "C_STRSHOWPAGING"
			NewColumns(3, TableName) = "CONFIG"
			NewColumns(3, DataType_Access) = "SMALLINT"
			NewColumns(3, DataType_SQL6) = "SMALLINT"
			NewColumns(3, DataType_SQL7) = "SMALLINT"
			NewColumns(3, DataType_MySQL) = "SMALLINT"
			NewColumns(3, ConstraintAccess)  = "NULL"
			NewColumns(3, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0365 DEFAULT 0"
			NewColumns(3, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0365 DEFAULT 0"
			NewColumns(3, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns(4, Prefix)	 = strTablePrefix
			NewColumns(4, FieldName) = "C_STRSHOWTOPICNAV"
			NewColumns(4, TableName) = "CONFIG"
			NewColumns(4, DataType_Access) = "SMALLINT"
			NewColumns(4, DataType_SQL6) = "SMALLINT"
			NewColumns(4, DataType_SQL7) = "SMALLINT"
			NewColumns(4, DataType_MySQL) = "SMALLINT"
			NewColumns(4, ConstraintAccess)  = "NULL"
			NewColumns(4, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0366 DEFAULT 0"
			NewColumns(4, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0366 DEFAULT 0"
			NewColumns(4, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns(5, Prefix)	 = strTablePrefix
			NewColumns(5, FieldName) = "C_STRPAGESIZE"
			NewColumns(5, TableName) = "CONFIG"
			NewColumns(5, DataType_Access) = "SMALLINT"
			NewColumns(5, DataType_SQL6) = "SMALLINT"
			NewColumns(5, DataType_SQL7) = "SMALLINT"
			NewColumns(5, DataType_MySQL) = "SMALLINT"
			NewColumns(5, ConstraintAccess)  = "NULL"
			NewColumns(5, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0367 DEFAULT 15"
			NewColumns(5, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0367 DEFAULT 15"
			NewColumns(5, ConstraintMySQL)  = "DEFAULT 15 NULL"

			NewColumns(6, Prefix)	 = strTablePrefix
			NewColumns(6, FieldName) = "C_STRPAGENUMBERSIZE"
			NewColumns(6, TableName) = "CONFIG"
			NewColumns(6, DataType_Access) = "SMALLINT"
			NewColumns(6, DataType_SQL6) = "SMALLINT"
			NewColumns(6, DataType_SQL7) = "SMALLINT"
			NewColumns(6, DataType_MySQL) = "SMALLINT"
			NewColumns(6, ConstraintAccess)  = "NULL"
			NewColumns(6, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0368 DEFAULT 10"
			NewColumns(6, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0368 DEFAULT 10"
			NewColumns(6, ConstraintMySQL)  = "DEFAULT 10 NULL"

			NewColumns(7, Prefix)	 = strTablePrefix
			NewColumns(7, FieldName) = "F_LAST_POST_AUTHOR"
			NewColumns(7, TableName) = "FORUM"
			NewColumns(7, DataType_Access) = "INT"
			NewColumns(7, DataType_SQL6) = "INT"
			NewColumns(7, DataType_SQL7) = "INT"
			NewColumns(7, DataType_MySQL) = "INT"
			NewColumns(7, ConstraintAccess)  = "NULL"
			NewColumns(7, ConstraintSQL6)  = "NULL"
			NewColumns(7, ConstraintSQL7)  = "NULL"
			NewColumns(7, ConstraintMySQL)  = "NULL"

			NewColumns(8, Prefix)	 = strTablePrefix
			NewColumns(8, FieldName) = "T_LAST_POST_AUTHOR"
			NewColumns(8, TableName) = "TOPICS"
			NewColumns(8, DataType)  = "INT"
			NewColumns(8, DataType_Access) = "INT"
			NewColumns(8, DataType_SQL6) = "INT"
			NewColumns(8, DataType_SQL7) = "INT"
			NewColumns(8, DataType_MySQL) = "INT"
			NewColumns(8, ConstraintAccess)  = "NULL"
			NewColumns(8, ConstraintSQL6)  = "NULL"
			NewColumns(8, ConstraintSQL7)  = "NULL"
			NewColumns(8, ConstraintMySQL)  = "NULL"

			call AddColumns(NewColumns, intCriticalErrors, intWarnings)

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRSHOWSTATISTICS        =  " & 1
			strSql = strSql & " ,    C_STRSHOWIMAGEPOWEREDBY    =  " & 1
			strSql = strSql & " ,    C_STRLOGONFORMAIL          =  " & 1
			strSql = strSql & " ,    C_STRSHOWPAGING            =  " & 0
			strSql = strSql & " ,    C_STRSHOWTOPICNAV          =  " & 0
			strSql = strSql & " ,    C_STRPAGESIZE              =  " & 15
			strSql = strSql & " ,    C_STRPAGENUMBERSIZE        =  " & 10
			strSql = strSql & " ,    C_STRVERSION               =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE CONFIG_ID = " & 1

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if
		if (OldVersion <= 2) then

			Dim NewColumns2(29,11)

			NewColumns2(0, Prefix)	 = strMemberTablePrefix
			NewColumns2(0, FieldName) = "M_FIRSTNAME"
			NewColumns2(0, TableName) = "MEMBERS"
			NewColumns2(0, DataType_Access)  = "TEXT (100)"
			NewColumns2(0, DataType_SQL6)  = "VARCHAR (100)"
			NewColumns2(0, DataType_SQL7)  = "NVARCHAR (100)"
			NewColumns2(0, DataType_MYSQL)  = "VARCHAR (100)"
			NewColumns2(0, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0369 DEFAULT ''"
			NewColumns2(0, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0369 DEFAULT ''"
			NewColumns2(0, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(1, Prefix)	 = strMemberTablePrefix
			NewColumns2(1, FieldName) = "M_LASTNAME"
			NewColumns2(1, TableName) = "MEMBERS"
			NewColumns2(1, DataType_Access)  = "TEXT (100)"
			NewColumns2(1, DataType_SQL6)  = "VARCHAR (100)"
			NewColumns2(1, DataType_SQL7)  = "NVARCHAR (100)"
			NewColumns2(1, DataType_MYSQL)  = "VARCHAR (100)"
			NewColumns2(1, ConstraintAccess)  = "NULL"
			NewColumns2(1, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0370 DEFAULT ''"
			NewColumns2(1, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0370 DEFAULT ''"
			NewColumns2(1, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(2, Prefix)	 = strMemberTablePrefix
			NewColumns2(2, FieldName) = "M_OCCUPATION"
			NewColumns2(2, TableName) = "MEMBERS"
			NewColumns2(2, DataType_Access)  = "TEXT (255)"
			NewColumns2(2, DataType_SQL6)  = "VARCHAR (255)"
			NewColumns2(2, DataType_SQL7)  = "NVARCHAR (255)"
			NewColumns2(2, DataType_MYSQL)  = "VARCHAR (255)"
			NewColumns2(2, ConstraintAccess)  = "NULL"
			NewColumns2(2, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0371 DEFAULT ''"
			NewColumns2(2, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0371 DEFAULT ''"
			NewColumns2(2, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(3, Prefix)	 = strMemberTablePrefix
			NewColumns2(3, FieldName) = "M_SEX"
			NewColumns2(3, TableName) = "MEMBERS"
			NewColumns2(3, DataType_Access)  = "TEXT (50)"
			NewColumns2(3, DataType_SQL6)  = "VARCHAR (50)"
			NewColumns2(3, DataType_SQL7)  = "NVARCHAR (50)"
			NewColumns2(3, DataType_MYSQL)  = "VARCHAR (50)"
			NewColumns2(3, ConstraintAccess)  = "NULL"
			NewColumns2(3, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0372 DEFAULT ''"
			NewColumns2(3, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0372 DEFAULT ''"
			NewColumns2(3, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(4, Prefix)	 = strMemberTablePrefix
			NewColumns2(4, FieldName) = "M_AGE"
			NewColumns2(4, TableName) = "MEMBERS"
			NewColumns2(4, DataType_Access)  = "TEXT (10)"
			NewColumns2(4, DataType_SQL6)  = "VARCHAR (10)"
			NewColumns2(4, DataType_SQL7)  = "NVARCHAR (10)"
			NewColumns2(4, DataType_MYSQL)  = "VARCHAR (10)"
			NewColumns2(4, ConstraintAccess)  = "NULL"
			NewColumns2(4, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0373 DEFAULT ''"
			NewColumns2(4, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0373 DEFAULT ''"
			NewColumns2(4, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(5, Prefix)	 = strMemberTablePrefix
			NewColumns2(5, FieldName) = "M_HOBBIES"
			NewColumns2(5, TableName) = "MEMBERS"
			NewColumns2(5, DataType_Access)  = "MEMO"
			NewColumns2(5, DataType_SQL6)  = "TEXT"
			NewColumns2(5, DataType_SQL7)  = "NTEXT"
			NewColumns2(5, DataType_MYSQL)  = "TEXT"
			NewColumns2(5, ConstraintAccess)  = "NULL"
			NewColumns2(5, ConstraintSQL6)  = "NULL"
			NewColumns2(5, ConstraintSQL7)  = "NULL"
			NewColumns2(5, ConstraintMySQL)  = "NULL"

			NewColumns2(6, Prefix)	 = strMemberTablePrefix
			NewColumns2(6, FieldName) = "M_LNEWS"
			NewColumns2(6, TableName) = "MEMBERS"
			NewColumns2(6, DataType_Access)  = "MEMO"
			NewColumns2(6, DataType_SQL6)  = "TEXT"
			NewColumns2(6, DataType_SQL7)  = "NTEXT"
			NewColumns2(6, DataType_MYSQL)  = "TEXT"
			NewColumns2(6, ConstraintAccess)  = "NULL"
			NewColumns2(6, ConstraintSQL6)  = "NULL"
			NewColumns2(6, ConstraintSQL7)  = "NULL"
			NewColumns2(6, ConstraintMySQL)  = "NULL"

			NewColumns2(7, Prefix)	 = strMemberTablePrefix
			NewColumns2(7, FieldName) = "M_QUOTE"
			NewColumns2(7, TableName) = "MEMBERS"
			NewColumns2(7, DataType_Access)  = "MEMO"
			NewColumns2(7, DataType_SQL6)  = "TEXT"
			NewColumns2(7, DataType_SQL7)  = "NTEXT"
			NewColumns2(7, DataType_MYSQL)  = "TEXT"
			NewColumns2(7, ConstraintAccess)  = "NULL"
			NewColumns2(7, ConstraintSQL6)  = "NULL"
			NewColumns2(7, ConstraintSQL7)  = "NULL"
			NewColumns2(7, ConstraintMySQL)  = "NULL"

			NewColumns2(8, Prefix)	 = strMemberTablePrefix
			NewColumns2(8, FieldName) = "M_BIO"
			NewColumns2(8, TableName) = "MEMBERS"
			NewColumns2(8, DataType_Access)  = "MEMO"
			NewColumns2(8, DataType_SQL6)  = "TEXT"
			NewColumns2(8, DataType_SQL7)  = "NTEXT"
			NewColumns2(8, DataType_MYSQL)  = "TEXT"
			NewColumns2(8, ConstraintAccess)  = "NULL"
			NewColumns2(8, ConstraintSQL6)  = "NULL"
			NewColumns2(8, ConstraintSQL7)  = "NULL"
			NewColumns2(8, ConstraintMySQL)  = "NULL"

			NewColumns2(9, Prefix)	 = strMemberTablePrefix
			NewColumns2(9, FieldName) = "M_MARSTATUS"
			NewColumns2(9, TableName) = "MEMBERS"
			NewColumns2(9, DataType_Access)  = "TEXT (100)"
			NewColumns2(9, DataType_SQL6)  = "VARCHAR (100)"
			NewColumns2(9, DataType_SQL7)  = "NVARCHAR (100)"
			NewColumns2(9, DataType_MYSQL)  = "VARCHAR (100)"
			NewColumns2(9, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0374 DEFAULT ''"
			NewColumns2(9, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0374 DEFAULT ''"
			NewColumns2(9, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(10, Prefix)	 = strMemberTablePrefix
			NewColumns2(10, FieldName) = "M_LINK1"
			NewColumns2(10, TableName) = "MEMBERS"
			NewColumns2(10, DataType_Access)  = "TEXT (255)"
			NewColumns2(10, DataType_SQL6)  = "VARCHAR (255)"
			NewColumns2(10, DataType_SQL7)  = "NVARCHAR (255)"
			NewColumns2(10, DataType_MYSQL)  = "VARCHAR (255)"
			NewColumns2(10, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0375 DEFAULT ''"
			NewColumns2(10, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0375 DEFAULT ''"
			NewColumns2(10, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(11, Prefix)	 = strMemberTablePrefix
			NewColumns2(11, FieldName) = "M_LINK2"
			NewColumns2(11, TableName) = "MEMBERS"
			NewColumns2(11, DataType_Access)  = "TEXT (255)"
			NewColumns2(11, DataType_SQL6)  = "VARCHAR (255)"
			NewColumns2(11, DataType_SQL7)  = "NVARCHAR (255)"
			NewColumns2(11, DataType_MYSQL)  = "VARCHAR (255)"
			NewColumns2(11, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0376 DEFAULT ''"
			NewColumns2(11, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0376 DEFAULT ''"
			NewColumns2(11, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(12, Prefix)	 = strMemberTablePrefix
			NewColumns2(12, FieldName) = "M_CITY"
			NewColumns2(12, TableName) = "MEMBERS"
			NewColumns2(12, DataType_Access)  = "TEXT (100)"
			NewColumns2(12, DataType_SQL6)  = "VARCHAR (100)"
			NewColumns2(12, DataType_SQL7)  = "NVARCHAR (100)"
			NewColumns2(12, DataType_MYSQL)  = "VARCHAR (100)"
			NewColumns2(12, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0377 DEFAULT ''"
			NewColumns2(12, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0377 DEFAULT ''"
			NewColumns2(12, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(13, Prefix)	 = strMemberTablePrefix
			NewColumns2(13, FieldName) = "M_PHOTO_URL"
			NewColumns2(13, TableName) = "MEMBERS"
			NewColumns2(13, DataType_Access)  = "TEXT (255)"
			NewColumns2(13, DataType_SQL6)  = "VARCHAR (255)"
			NewColumns2(13, DataType_SQL7)  = "NVARCHAR (255)"
			NewColumns2(13, DataType_MYSQL)  = "VARCHAR (255)"
			NewColumns2(13, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0378 DEFAULT ''"
			NewColumns2(13, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0378 DEFAULT ''"
			NewColumns2(13, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(14, Prefix)	 = strMemberTablePrefix
			NewColumns2(14, FieldName) = "M_STATE"
			NewColumns2(14, TableName) = "MEMBERS"
			NewColumns2(14, DataType_Access)  = "TEXT (100)"
			NewColumns2(14, DataType_SQL6)  = "VARCHAR (100)"
			NewColumns2(14, DataType_SQL7)  = "NVARCHAR (100)"
			NewColumns2(14, DataType_MYSQL)  = "VARCHAR (100)"
			NewColumns2(14, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0379 DEFAULT ''"
			NewColumns2(14, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0379 DEFAULT ''"
			NewColumns2(14, ConstraintMySQL)  = "DEFAULT '' NULL"

			NewColumns2(15, Prefix)	 = strTablePrefix
			NewColumns2(15, FieldName) = "C_STRFULLNAME"
			NewColumns2(15, TableName) = "CONFIG"
			NewColumns2(15, DataType_Access) = "SMALLINT"
			NewColumns2(15, DataType_SQL6) = "SMALLINT"
			NewColumns2(15, DataType_SQL7) = "SMALLINT"
			NewColumns2(15, DataType_MySQL) = "SMALLINT"
			NewColumns2(15, ConstraintAccess)  = "NULL"
			NewColumns2(15, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1000 DEFAULT 0"
			NewColumns2(15, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1000 DEFAULT 0"
			NewColumns2(15, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(16, Prefix)	 = strTablePrefix
			NewColumns2(16, FieldName) = "C_STRPICTURE"
			NewColumns2(16, TableName) = "CONFIG"
			NewColumns2(16, DataType_Access) = "SMALLINT"
			NewColumns2(16, DataType_SQL6) = "SMALLINT"
			NewColumns2(16, DataType_SQL7) = "SMALLINT"
			NewColumns2(16, DataType_MySQL) = "SMALLINT"
			NewColumns2(16, ConstraintAccess)  = "NULL"
			NewColumns2(16, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1001 DEFAULT 0"
			NewColumns2(16, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1001 DEFAULT 0"
			NewColumns2(16, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(17, Prefix)	 = strTablePrefix
			NewColumns2(17, FieldName) = "C_STRSEX"
			NewColumns2(17, TableName) = "CONFIG"
			NewColumns2(17, DataType_Access) = "SMALLINT"
			NewColumns2(17, DataType_SQL6) = "SMALLINT"
			NewColumns2(17, DataType_SQL7) = "SMALLINT"
			NewColumns2(17, DataType_MySQL) = "SMALLINT"
			NewColumns2(17, ConstraintAccess)  = "NULL"
			NewColumns2(17, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1002 DEFAULT 0"
			NewColumns2(17, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1002 DEFAULT 0"
			NewColumns2(17, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(18, Prefix)	 = strTablePrefix
			NewColumns2(18, FieldName) = "C_STRCITY"
			NewColumns2(18, TableName) = "CONFIG"
			NewColumns2(18, DataType_Access) = "SMALLINT"
			NewColumns2(18, DataType_SQL6) = "SMALLINT"
			NewColumns2(18, DataType_SQL7) = "SMALLINT"
			NewColumns2(18, DataType_MySQL) = "SMALLINT"
			NewColumns2(18, ConstraintAccess)  = "NULL"
			NewColumns2(18, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1003 DEFAULT 0"
			NewColumns2(18, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1003 DEFAULT 0"
			NewColumns2(18, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(19, Prefix)	 = strTablePrefix
			NewColumns2(19, FieldName) = "C_STRSTATE"
			NewColumns2(19, TableName) = "CONFIG"
			NewColumns2(19, DataType_Access) = "SMALLINT"
			NewColumns2(19, DataType_SQL6) = "SMALLINT"
			NewColumns2(19, DataType_SQL7) = "SMALLINT"
			NewColumns2(19, DataType_MySQL) = "SMALLINT"
			NewColumns2(19, ConstraintAccess)  = "NULL"
			NewColumns2(19, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1004 DEFAULT 0"
			NewColumns2(19, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1004 DEFAULT 0"
			NewColumns2(19, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(20, Prefix)	 = strTablePrefix
			NewColumns2(20, FieldName) = "C_STRAGE"
			NewColumns2(20, TableName) = "CONFIG"
			NewColumns2(20, DataType_Access) = "SMALLINT"
			NewColumns2(20, DataType_SQL6) = "SMALLINT"
			NewColumns2(20, DataType_SQL7) = "SMALLINT"
			NewColumns2(20, DataType_MySQL) = "SMALLINT"
			NewColumns2(20, ConstraintAccess)  = "NULL"
			NewColumns2(20, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1005 DEFAULT 0"
			NewColumns2(20, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1005 DEFAULT 0"
			NewColumns2(20, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(21, Prefix)	 = strTablePrefix
			NewColumns2(21, FieldName) = "C_STRCOUNTRY"
			NewColumns2(21, TableName) = "CONFIG"
			NewColumns2(21, DataType_Access) = "SMALLINT"
			NewColumns2(21, DataType_SQL6) = "SMALLINT"
			NewColumns2(21, DataType_SQL7) = "SMALLINT"
			NewColumns2(21, DataType_MySQL) = "SMALLINT"
			NewColumns2(21, ConstraintAccess)  = "NULL"
			NewColumns2(21, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1006 DEFAULT 0"
			NewColumns2(21, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1006 DEFAULT 0"
			NewColumns2(21, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(22, Prefix)	 = strTablePrefix
			NewColumns2(22, FieldName) = "C_STROCCUPATION"
			NewColumns2(22, TableName) = "CONFIG"
			NewColumns2(22, DataType_Access) = "SMALLINT"
			NewColumns2(22, DataType_SQL6) = "SMALLINT"
			NewColumns2(22, DataType_SQL7) = "SMALLINT"
			NewColumns2(22, DataType_MySQL) = "SMALLINT"
			NewColumns2(22, ConstraintAccess)  = "NULL"
			NewColumns2(22, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1007 DEFAULT 0"
			NewColumns2(22, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1007 DEFAULT 0"
			NewColumns2(22, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(23, Prefix)	 = strTablePrefix
			NewColumns2(23, FieldName) = "C_STRBIO"
			NewColumns2(23, TableName) = "CONFIG"
			NewColumns2(23, DataType_Access) = "SMALLINT"
			NewColumns2(23, DataType_SQL6) = "SMALLINT"
			NewColumns2(23, DataType_SQL7) = "SMALLINT"
			NewColumns2(23, DataType_MySQL) = "SMALLINT"
			NewColumns2(23, ConstraintAccess)  = "NULL"
			NewColumns2(23, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1008 DEFAULT 0"
			NewColumns2(23, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1008 DEFAULT 0"
			NewColumns2(23, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(24, Prefix)	 = strTablePrefix
			NewColumns2(24, FieldName) = "C_STRHOBBIES"
			NewColumns2(24, TableName) = "CONFIG"
			NewColumns2(24, DataType_Access) = "SMALLINT"
			NewColumns2(24, DataType_SQL6) = "SMALLINT"
			NewColumns2(24, DataType_SQL7) = "SMALLINT"
			NewColumns2(24, DataType_MySQL) = "SMALLINT"
			NewColumns2(24, ConstraintAccess)  = "NULL"
			NewColumns2(24, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1009 DEFAULT 0"
			NewColumns2(24, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1009 DEFAULT 0"
			NewColumns2(24, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(25, Prefix)	 = strTablePrefix
			NewColumns2(25, FieldName) = "C_STRLNEWS"
			NewColumns2(25, TableName) = "CONFIG"
			NewColumns2(25, DataType_Access) = "SMALLINT"
			NewColumns2(25, DataType_SQL6) = "SMALLINT"
			NewColumns2(25, DataType_SQL7) = "SMALLINT"
			NewColumns2(25, DataType_MySQL) = "SMALLINT"
			NewColumns2(25, ConstraintAccess)  = "NULL"
			NewColumns2(25, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1010 DEFAULT 0"
			NewColumns2(25, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1010 DEFAULT 0"
			NewColumns2(25, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(26, Prefix)	 = strTablePrefix
			NewColumns2(26, FieldName) = "C_STRQUOTE"
			NewColumns2(26, TableName) = "CONFIG"
			NewColumns2(26, DataType_Access) = "SMALLINT"
			NewColumns2(26, DataType_SQL6) = "SMALLINT"
			NewColumns2(26, DataType_SQL7) = "SMALLINT"
			NewColumns2(26, DataType_MySQL) = "SMALLINT"
			NewColumns2(26, ConstraintAccess)  = "NULL"
			NewColumns2(26, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1011 DEFAULT 0"
			NewColumns2(26, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1011 DEFAULT 0"
			NewColumns2(26, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(27, Prefix)	 = strTablePrefix
			NewColumns2(27, FieldName) = "C_STRMARSTATUS"
			NewColumns2(27, TableName) = "CONFIG"
			NewColumns2(27, DataType_Access) = "SMALLINT"
			NewColumns2(27, DataType_SQL6) = "SMALLINT"
			NewColumns2(27, DataType_SQL7) = "SMALLINT"
			NewColumns2(27, DataType_MySQL) = "SMALLINT"
			NewColumns2(27, ConstraintAccess)  = "NULL"
			NewColumns2(27, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1012 DEFAULT 0"
			NewColumns2(27, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1012 DEFAULT 0"
			NewColumns2(27, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(28, Prefix)	 = strTablePrefix
			NewColumns2(28, FieldName) = "C_STRFAVLINKS"
			NewColumns2(28, TableName) = "CONFIG"
			NewColumns2(28, DataType_Access) = "SMALLINT"
			NewColumns2(28, DataType_SQL6) = "SMALLINT"
			NewColumns2(28, DataType_SQL7) = "SMALLINT"
			NewColumns2(28, DataType_MySQL) = "SMALLINT"
			NewColumns2(28, ConstraintAccess)  = "NULL"
			NewColumns2(28, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1013 DEFAULT 0"
			NewColumns2(28, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1013 DEFAULT 0"
			NewColumns2(28, ConstraintMySQL)  = "DEFAULT 0 NULL"

			NewColumns2(29, Prefix)	 = strTablePrefix
			NewColumns2(29, FieldName) = "C_STRRECENTTOPICS"
			NewColumns2(29, TableName) = "CONFIG"
			NewColumns2(29, DataType_Access) = "SMALLINT"
			NewColumns2(29, DataType_SQL6) = "SMALLINT"
			NewColumns2(29, DataType_SQL7) = "SMALLINT"
			NewColumns2(29, DataType_MySQL) = "SMALLINT"
			NewColumns2(29, ConstraintAccess)  = "NULL"
			NewColumns2(29, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1014 DEFAULT 0"
			NewColumns2(29, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1014 DEFAULT 0"
			NewColumns2(29, ConstraintMySQL)  = "DEFAULT 0 NULL"

			call AddColumns(NewColumns2, intCriticalErrors, intWarnings)

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
			strSql = strSql & " , C_STRFULLNAME   =  " & 0
			strSql = strSql & " , C_STRPICTURE    =  " & 0
			strSql = strSql & " , C_STRSEX        =  " & 0
			strSql = strSql & " , C_STRCITY       =  " & 0
			strSql = strSql & " , C_STRSTATE      =  " & 0
			strSql = strSql & " , C_STRAGE        =  " & 0
			strSql = strSql & " , C_STRCOUNTRY    =  " & 1
			strSql = strSql & " , C_STROCCUPATION =  " & 0
			strSql = strSql & " , C_STRHOMEPAGE   =  " & 1
			strSql = strSql & " , C_STRFAVLINKS   =  " & 1
			strSql = strSql & " , C_STRICQ        =  " & 1
			strSql = strSql & " , C_STRYAHOO      =  " & 1
			strSql = strSql & " , C_STRAIM        =  " & 1
			strSql = strSql & " , C_STRBIO        =  " & 0
			strSql = strSql & " , C_STRHOBBIES    =  " & 0
			strSql = strSql & " , C_STRLNEWS      =  " & 0
			strSql = strSql & " , C_STRQUOTE      =  " & 0
			strSql = strSql & " , C_STRMARSTATUS  =  " & 0
			strSql = strSql & " , C_STRRECENTTOPICS  =  " & 0
			strSql = strSql & " WHERE CONFIG_ID = " & 1

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush

		end if

'###########################################################################
'##
'## Set up for update 3
'##
'## 	Database updates needed
'##
'## 	FORUM_FORUM
'## 	Need F_PASSWORD set to 255 Char's to handle NT Group Names.
'##
'###########################################################################

		if (OldVersion <= 3) then

			Dim NewColumns3(2,11)

			NewColumns3(0, Prefix)	 = strTablePrefix
			NewColumns3(0, FieldName) = "C_STRAUTOLOGON"
			NewColumns3(0, TableName) = "CONFIG"
			NewColumns3(0, DataType_Access) = "SMALLINT"
			NewColumns3(0, DataType_SQL6) = "SMALLINT"
			NewColumns3(0, DataType_SQL7) = "SMALLINT"
			NewColumns3(0, DataType_MySQL) = "SMALLINT"
			NewColumns3(0, ConstraintAccess)  = "NULL"
			NewColumns3(0, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1015 DEFAULT 0"
			NewColumns3(0, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1015 DEFAULT 0"
			NewColumns3(0, ConstraintMySQL)  = "DEFAULT '0' NULL"

			NewColumns3(1, Prefix)	 = strTablePrefix
			NewColumns3(1, FieldName) = "C_STRNTGROUPS"
			NewColumns3(1, TableName) = "CONFIG"
			NewColumns3(1, DataType_Access) = "SMALLINT"
			NewColumns3(1, DataType_SQL6) = "SMALLINT"
			NewColumns3(1, DataType_SQL7) = "SMALLINT"
			NewColumns3(1, DataType_MySQL) = "SMALLINT"
			NewColumns3(1, ConstraintAccess)  = "NULL"
			NewColumns3(1, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1016 DEFAULT 0"
			NewColumns3(1, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1016 DEFAULT 0"
			NewColumns3(1, ConstraintMySQL)  = "DEFAULT '0' NULL"

			NewColumns3(2, Prefix)	 = strTablePrefix
			NewColumns3(2, FieldName) = "F_PASSWORD_NEW"
			NewColumns3(2, TableName) = "FORUM"
			NewColumns3(2, DataType_Access)  = "TEXT (255)"
			NewColumns3(2, DataType_SQL6)  = "VARCHAR (255)"
			NewColumns3(2, DataType_SQL7)  = "NVARCHAR (255)"
			NewColumns3(2, DataType_MYSQL)  = "VARCHAR (255)"
			NewColumns3(2, ConstraintAccess)  = "NULL"
			NewColumns3(2, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT ''"
			NewColumns3(2, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT ''"
			NewColumns3(2, ConstraintMySQL)  = "DEFAULT '' NULL"

			call AddColumns(NewColumns3, intCriticalErrors, intWarnings)

			Dim SpecialSql3(4)

			SpecialSql3(Access) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
			SpecialSql3(SQL6) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
			SpecialSql3(SQL7) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
			SpecialSql3(MySql) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
			strOkMessage = "Password field conversion step 1 of 2 completed"

			call SpecialUpdates(SpecialSql3, strOkMessage)

			SpecialSql3(Access) = "ALTER TABLE " & strTablePrefix & "FORUM DROP COLUMN F_PASSWORD"
			SpecialSql3(SQL6) = "SELECT * FROM " & strTablePrefix & "CONFIG " '## dummy sql-statement SQL6.5 doesn't allow DROP !!
			SpecialSql3(SQL7) = "ALTER TABLE " & strTablePrefix & "FORUM DROP COLUMN F_PASSWORD"
			SpecialSql3(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM DROP COLUMN F_PASSWORD"
			strOkMessage = "Password field conversion step 2 of 2 completed"

			call SpecialUpdates(SpecialSql3, strOkMessage)

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
			strSql = strSql & " , C_STRAUTOLOGON  =  " & 0
			strSql = strSql & " , C_STRNTGROUPS   =  " & 0
			strSql = strSql & " WHERE CONFIG_ID  =  " & 1

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush

		end if

'###########################################################################
'##
'## Setup for update 4
'##
'###########################################################################

		if (OldVersion <= 4) then

			Dim SpecialSql4(4)

			SpecialSql4(Access) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
 			SpecialSql4(Access) = SpecialSql4(Access) & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL, "
 			SpecialSql4(Access) = SpecialSql4(Access) & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY (MEMBER_ID, FORUM_ID) ) "

 			SpecialSql4(SQL6) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
			SpecialSql4(SQL6) = SpecialSql4(SQL6) & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL , "
			SpecialSql4(SQL6) = SpecialSql4(SQL6) & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY NONCLUSTERED (MEMBER_ID, FORUM_ID) )"

			SpecialSql4(SQL7) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
			SpecialSql4(SQL7) = SpecialSql4(SQL7) & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL , "
			SpecialSql4(SQL7) = SpecialSql4(SQL7) & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY NONCLUSTERED (MEMBER_ID, FORUM_ID) )"

			SpecialSql4(MySql) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
			SpecialSql4(MySql) = SpecialSql4(MySql) & "MEMBER_ID INT (11) NOT NULL, FORUM_ID smallint (6) NOT NULL , "
			SpecialSql4(MySql) = SpecialSql4(MySql) & "PRIMARY KEY (MEMBER_ID, FORUM_ID) ) "

	 		strOkMessage = "Table ALLOWED_MEMBERS created "

			call SpecialUpdates(SpecialSql4, strOkMessage)

			Response.Flush

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)
			Response.Write("  <tr>" & vbNewLine)
			Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgrading: </b></font></td>" & vbNewLine)
			Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Starting transferring Member List to ALLOWED_MEMBERS table</font></td>" & vbNewLine)
			Response.Write("  </tr>" & vbNewLine)

			intTransferErrors = 0
			strSql = "SELECT FORUM_ID,F_USERLIST FROM " & strTablePrefix & "FORUM "

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear

			set rsForum = my_Conn.execute(strSql)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Table opened</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Error while getting Memberlist for transfer<b></font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
				intTransferErrors = 1
			end if

			if intTransferErrors = 0 then
				do while not rsForum.EOF
					if Instr(rsForum("F_USERLIST"),",") > 0 then
						Users = split(rsForum("F_USERLIST"),",")
							for count = Lbound(Users) to Ubound(Users)
								strSql = "INSERT INTO " & strTablePrefix & "ALLOWED_MEMBERS ("
								strSql = strSql & " MEMBER_ID, FORUM_ID) VALUES ( "& Users(count) & ", " & rsForum("FORUM_ID") & ")"

								on error resume next
								my_Conn.Errors.Clear
								Err.Clear
								my_Conn.Execute (strSql)

								UpdateErrorCode = UpdateErrorCheck()

								on error goto 0
								if UpdateErrorCode = 0 then
									Response.Write("  <tr>" & vbNewLine)
									Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
									Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
									Response.Write("  </tr>" & vbNewLine)
								elseif UpdateErrorCode = 2 then
									Response.Write("  <tr>" & vbNewLine)
									Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
									Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Error while adding record ( " & Users(count) & ", " & rsForum("FORUM_ID") & ") to ALLOWED_MEMBERS table!<b></font></td>" & vbNewLine)
									Response.Write("  </tr>" & vbNewLine)
									intCriticalErrors = intCriticalErrors + 1
									infTransferErrors = 1
								end if
							next
					end if
		 			rsForum.movenext
				loop
			end if
			on error resume next
			rsForum.close
			set rsForum = nothing
			on error goto 0

			if intTransferErrors = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgrading: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Finished transferring Member List to ALLOWED_MEMBERS table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgrading: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Transferring of Member List to ALLOWED_MEMBERS table was NOT succesfull !</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			end if
			Response.Write("  </table>" & vbNewLine)

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE CONFIG_ID  =  " & 1

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if

'###########################################################################
'##
'## Setup for update 5 / to version 3.3
'##
'###########################################################################

		if (OldVersion <= 5) then

			Dim SpecialSql5(4)

			if strDBType = "access" then
				SpecialSql5(Access) = "CREATE TABLE " & strTablePrefix & "CONFIG_NEW ( "
 				SpecialSql5(Access) = SpecialSql5(Access) & "ID COUNTER NOT NULL , "
	 			SpecialSql5(Access) = SpecialSql5(Access) & "C_VARIABLE varchar (255) NULL , "
	 			SpecialSql5(Access) = SpecialSql5(Access) & "C_VALUE varchar (255) NULL )"

		 		strOkMessage = "Table CONFIG_NEW created "

				call SpecialUpdates(SpecialSql5, strOkMessage)
			end if

			SpecialSql5(Access) = "CREATE TABLE " & strTablePrefix & "SUBSCRIPTIONS ("
 			SpecialSql5(Access) = SpecialSql5(Access) & "SUBSCRIPTION_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, MEMBER_ID INT NOT NULL, "
 			SpecialSql5(Access) = SpecialSql5(Access) & "CAT_ID INT NOT NULL, TOPIC_ID INT NOT NULL, FORUM_ID INT NOT NULL) "

			SpecialSql5(SQL6) = "CREATE TABLE " & strTablePrefix & "SUBSCRIPTIONS ("
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "SUBSCRIPTION_ID INT IDENTITY NOT NULL, MEMBER_ID INT NOT NULL, "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "CAT_ID INT NOT NULL, TOPIC_ID INT NOT NULL, FORUM_ID INT NOT NULL) "

			SpecialSql5(SQL7) = "CREATE TABLE " & strTablePrefix & "SUBSCRIPTIONS ("
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "SUBSCRIPTION_ID INT IDENTITY NOT NULL, MEMBER_ID INT NOT NULL, "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "CAT_ID INT NOT NULL, TOPIC_ID INT NOT NULL, FORUM_ID INT NOT NULL) "

			SpecialSql5(MySql) = "CREATE TABLE " & strTablePrefix & "SUBSCRIPTIONS ("
 			SpecialSql5(MySql) = SpecialSql5(MySql) & "SUBSCRIPTION_ID INT (11) NOT NULL auto_increment, MEMBER_ID INT NOT NULL, "
 			SpecialSql5(MySql) = SpecialSql5(MySql) & "CAT_ID INT NOT NULL, TOPIC_ID INT NOT NULL, FORUM_ID INT NOT NULL, "
	 		SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "SUBSCRIPTIONS_SUB_ID(SUBSCRIPTION_ID)) "

	 		strOkMessage = "Table SUBSCRIPTIONS created "

			call SpecialUpdates(SpecialSql5, strOkMessage)

			Response.Flush

			SpecialSql5(Access) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 			SpecialSql5(Access) = SpecialSql5(Access) & "CAT_ID int NOT NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "FORUM_ID int NOT NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "TOPIC_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_STATUS smallint NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_MAIL smallint NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "T_SUBJECT varchar (100) NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_MESSAGE text NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_AUTHOR int NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "T_REPLIES int NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_VIEW_COUNT int NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_LAST_POST varchar (14) NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "T_DATE varchar (14) NULL, "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_LAST_POSTER int NULL, "
 			SpecialSql5(Access) = SpecialSql5(Access) & "T_IP varchar (15) NULL, " 
	 		SpecialSql5(Access) = SpecialSql5(Access) & "T_LAST_POST_AUTHOR int NULL ) "

			SpecialSql5(SQL6) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "CAT_ID int NOT NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "FORUM_ID int NOT NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "TOPIC_ID int IDENTITY (1, 1) NOT NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_STATUS smallint NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_MAIL smallint NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_SUBJECT varchar (100) NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_MESSAGE text NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_AUTHOR int NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_REPLIES int NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_VIEW_COUNT int NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_LAST_POST varchar (14) NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_DATE varchar (14) NULL, "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_LAST_POSTER int NULL, "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_IP varchar (15) NULL, " 
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "T_LAST_POST_AUTHOR int NULL ) "

			SpecialSql5(SQL7) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "CAT_ID int NOT NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "FORUM_ID int NOT NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "TOPIC_ID int IDENTITY (1, 1) NOT NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_STATUS smallint NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_MAIL smallint NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_SUBJECT nvarchar (100) NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_MESSAGE ntext NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_AUTHOR int NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_REPLIES int NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_VIEW_COUNT int NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_LAST_POST varchar (14) NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_DATE varchar (14) NULL, "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_LAST_POSTER int NULL, "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_IP varchar (15) NULL, " 
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "T_LAST_POST_AUTHOR int NULL ) "

	 		SpecialSql5(MySql) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 			SpecialSql5(MySql) = SpecialSql5(MySql) & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
 			SpecialSql5(MySql) = SpecialSql5(MySql) & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "TOPIC_ID int (11) NOT NULL auto_increment, "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_STATUS smallint (6) DEFAULT '1' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_MAIL smallint (6) DEFAULT '0' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_SUBJECT VARCHAR (100) DEFAULT '' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_MESSAGE text , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_AUTHOR int (11) DEFAULT '1' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_REPLIES int (11) DEFAULT '0' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_VIEW_COUNT int (11) DEFAULT '0' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_LAST_POST VARCHAR (14) DEFAULT '' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_DATE VARCHAR (14) DEFAULT '', "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_LAST_POSTER int (11) DEFAULT '1', "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_IP VARCHAR (15) DEFAULT '000.000.000.000', " 
			SpecialSql5(MySql) = SpecialSql5(MySql) & "T_LAST_POST_AUTHOR int (11) DEFAULT '1',   "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_TOPIC_CATFORTOP(CAT_ID,FORUM_ID,TOPIC_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_TOPIC_CAT_ID(CAT_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_TOPIC_FORUM_ID(FORUM_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_TOPIC_TOPIC_ID (TOPIC_ID) )"

	 		strOkMessage = "Table A_TOPICS created "

			call SpecialUpdates(SpecialSql5, strOkMessage)

			Response.Flush

	 		SpecialSql5(Access) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
 			SpecialSql5(Access) = SpecialSql5(Access) & "CAT_ID int NOT NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "FORUM_ID int NOT NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "TOPIC_ID int NOT NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "REPLY_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "R_STATUS smallint NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "R_MAIL smallint NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "R_AUTHOR int NULL , "
	 		SpecialSql5(Access) = SpecialSql5(Access) & "R_MESSAGE text NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "R_DATE varchar (14) NULL , "
 			SpecialSql5(Access) = SpecialSql5(Access) & "R_IP varchar (15) NULL ) "

	 		SpecialSql5(SQL6) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "CAT_ID int NOT NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "FORUM_ID int NOT NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "TOPIC_ID int NOT NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "REPLY_ID int IDENTITY (1, 1) NOT NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "R_MAIL smallint NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "R_STATUS smallint NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "R_AUTHOR int NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "R_MESSAGE text NULL , "
	 		SpecialSql5(SQL6) = SpecialSql5(SQL6) & "R_DATE varchar (14) NULL , "
 			SpecialSql5(SQL6) = SpecialSql5(SQL6) & "R_IP varchar (15) NULL ) "

	 		SpecialSql5(SQL7) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "CAT_ID int NOT NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "FORUM_ID int NOT NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "TOPIC_ID int NOT NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "REPLY_ID int IDENTITY (1, 1) NOT NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "R_STATUS smallint NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "R_MAIL smallint NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "R_AUTHOR int NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "R_MESSAGE ntext NULL , "
	 		SpecialSql5(SQL7) = SpecialSql5(SQL7) & "R_DATE varchar (14) NULL , "
 			SpecialSql5(SQL7) = SpecialSql5(SQL7) & "R_IP varchar (15) NULL ) "

			SpecialSql5(MySql) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "TOPIC_ID int (11) DEFAULT '1' NOT NULL , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "REPLY_ID int (11) NOT NULL auto_increment, "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "R_STATUS smallint (6) DEFAULT '1' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "R_AUTHOR int (11) DEFAULT '1' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "R_MESSAGE text , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "R_DATE VARCHAR (14) DEFAULT '' , "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "R_IP VARCHAR (15) DEFAULT '000.000.000.000', "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_REPLY_CATFORTOPREPL(CAT_ID,FORUM_ID,TOPIC_ID, REPLY_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_REPLY_REP_ID(REPLY_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_REPLY_CAT_ID(CAT_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_REPLY_FORUM_ID(FORUM_ID), "
			SpecialSql5(MySql) = SpecialSql5(MySql) & "KEY " & strTablePrefix & "A_REPLY_TOPIC_ID (TOPIC_ID) )"

	 		strOkMessage = "Table A_REPLY created "

			call SpecialUpdates(SpecialSql5, strOkMessage)

			Response.Flush

			Dim NewColumns5(11,11)

			NewColumns5(0, Prefix)	 = strTablePrefix
			NewColumns5(0, FieldName) = "R_STATUS"
			NewColumns5(0, TableName) = "REPLY"
			NewColumns5(0, DataType_Access) = "SMALLINT"
			NewColumns5(0, DataType_SQL6) = "SMALLINT"
			NewColumns5(0, DataType_SQL7) = "SMALLINT"
			NewColumns5(0, DataType_MySQL) = "SMALLINT"
			NewColumns5(0, ConstraintAccess)  = "NOT NULL"
			NewColumns5(0, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT 0"
			NewColumns5(0, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT 0"
			NewColumns5(0, ConstraintMySQL)  = "DEFAULT '0' NOT NULL"

			NewColumns5(1, Prefix)	 = strTablePrefix
			NewColumns5(1, FieldName) = "F_MODERATION"
			NewColumns5(1, TableName) = "FORUM"
			NewColumns5(1, DataType_Access) = "INT"
			NewColumns5(1, DataType_SQL6) = "INT"
			NewColumns5(1, DataType_SQL7) = "INT"
			NewColumns5(1, DataType_MySQL) = "INT"
			NewColumns5(1, ConstraintAccess)  = "NOT NULL"
			NewColumns5(1, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1018 DEFAULT 0"
			NewColumns5(1, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1018 DEFAULT 0"
			NewColumns5(1, ConstraintMySQL)  = "DEFAULT '0' NOT NULL"

			NewColumns5(2, Prefix)	 = strTablePrefix
			NewColumns5(2, FieldName) = "F_SUBSCRIPTION"
			NewColumns5(2, TableName) = "FORUM"
			NewColumns5(2, DataType_Access) = "INT"
			NewColumns5(2, DataType_SQL6) = "INT"
			NewColumns5(2, DataType_SQL7) = "INT"
			NewColumns5(2, DataType_MySQL) = "INT"
			NewColumns5(2, ConstraintAccess)  = "NOT NULL"
			NewColumns5(2, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1019 DEFAULT 0"
			NewColumns5(2, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1019 DEFAULT 0"
			NewColumns5(2, ConstraintMySQL)  = "DEFAULT '0' NOT NULL"

			NewColumns5(3, Prefix)	 = strTablePrefix
			NewColumns5(3, FieldName) = "F_ORDER"
			NewColumns5(3, TableName) = "FORUM"
			NewColumns5(3, DataType_Access) = "INT"
			NewColumns5(3, DataType_SQL6) = "INT"
			NewColumns5(3, DataType_SQL7) = "INT"
			NewColumns5(3, DataType_MySQL) = "INT"
			NewColumns5(3, ConstraintAccess)  = "NOT NULL"
			NewColumns5(3, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1031 DEFAULT 1"
			NewColumns5(3, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1031 DEFAULT 1"
			NewColumns5(3, ConstraintMySQL)  = "DEFAULT '1' NOT NULL"

			NewColumns5(4, Prefix)	 = strTablePrefix
			NewColumns5(4, FieldName) = "CAT_MODERATION"
			NewColumns5(4, TableName) = "CATEGORY"
			NewColumns5(4, DataType_Access) = "SMALLINT"
			NewColumns5(4, DataType_SQL6) = "SMALLINT"
			NewColumns5(4, DataType_SQL7) = "SMALLINT"
			NewColumns5(4, DataType_MySQL) = "SMALLINT"
			NewColumns5(4, ConstraintAccess)  = "NOT NULL"
			NewColumns5(4, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1021 DEFAULT 0"
			NewColumns5(4, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1021 DEFAULT 0"
			NewColumns5(4, ConstraintMySQL)  = "DEFAULT '0' NOT NULL"

			NewColumns5(5, Prefix)	 = strTablePrefix
			NewColumns5(5, FieldName) = "CAT_SUBSCRIPTION"
			NewColumns5(5, TableName) = "CATEGORY"
			NewColumns5(5, DataType_Access) = "SMALLINT"
			NewColumns5(5, DataType_SQL6) = "SMALLINT"
			NewColumns5(5, DataType_SQL7) = "SMALLINT"
			NewColumns5(5, DataType_MySQL) = "SMALLINT"
			NewColumns5(5, ConstraintAccess)  = "NOT NULL"
			NewColumns5(5, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1022 DEFAULT 0"
			NewColumns5(5, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1022 DEFAULT 0"
			NewColumns5(5, ConstraintMySQL)  = "DEFAULT '0' NOT NULL"

			NewColumns5(6, Prefix)	 = strTablePrefix
			NewColumns5(6, FieldName) = "CAT_ORDER"
			NewColumns5(6, TableName) = "CATEGORY"
			NewColumns5(6, DataType_Access) = "SMALLINT"
			NewColumns5(6, DataType_SQL6) = "SMALLINT"
			NewColumns5(6, DataType_SQL7) = "SMALLINT"
			NewColumns5(6, DataType_MySQL) = "SMALLINT"
			NewColumns5(6, ConstraintAccess)  = "NOT NULL"
			NewColumns5(6, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1023 DEFAULT 1"
			NewColumns5(6, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1023 DEFAULT 1"
			NewColumns5(6, ConstraintMySQL)  = "DEFAULT '1' NOT NULL"

			NewColumns5(7, Prefix)	 = strTablePrefix
			NewColumns5(7, FieldName) = "T_ARCHIVE_FLAG"
			NewColumns5(7, TableName) = "TOPICS"
			NewColumns5(7, DataType_Access) = "SMALLINT"
			NewColumns5(7, DataType_SQL6) = "SMALLINT"
			NewColumns5(7, DataType_SQL7) = "SMALLINT"
			NewColumns5(7, DataType_MySQL) = "SMALLINT"
			NewColumns5(7, ConstraintAccess)  = "NOT NULL"
			NewColumns5(7, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1029 DEFAULT 1"
			NewColumns5(7, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1029 DEFAULT 1"
			NewColumns5(7, ConstraintMySQL)  = "DEFAULT '1' NOT NULL"

			NewColumns5(8, Prefix)	 = strTablePrefix
			NewColumns5(8, FieldName) = "F_L_ARCHIVE"
			NewColumns5(8, TableName) = "FORUM"
			NewColumns5(8, DataType_Access)  = "TEXT (14)"
			NewColumns5(8, DataType_SQL6)  = "VARCHAR (14)"
			NewColumns5(8, DataType_SQL7)  = "NVARCHAR (14)"
			NewColumns5(8, DataType_MYSQL)  = "VARCHAR (14)"
			NewColumns5(8, ConstraintAccess)  = "NOT NULL"
			NewColumns5(8, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1024 DEFAULT ''"
			NewColumns5(8, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1024 DEFAULT ''"
			NewColumns5(8, ConstraintMySQL)  = "DEFAULT '' NOT NULL"

			NewColumns5(9, Prefix)	 = strTablePrefix
			NewColumns5(9, FieldName) = "F_ARCHIVE_SCHED"
			NewColumns5(9, TableName) = "FORUM"
			NewColumns5(9, DataType_Access) = "INT"
			NewColumns5(9, DataType_SQL6) = "INT"
			NewColumns5(9, DataType_SQL7) = "INT"
			NewColumns5(9, DataType_MySQL) = "INT"
			NewColumns5(9, ConstraintAccess)  = "NOT NULL"
			NewColumns5(9, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1025 DEFAULT 30"
			NewColumns5(9, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1025 DEFAULT 30"
			NewColumns5(9, ConstraintMySQL)  = "DEFAULT '30' NOT NULL"

			NewColumns5(10, Prefix)	 = strTablePrefix
			NewColumns5(10, FieldName) = "F_L_DELETE"
			NewColumns5(10, TableName) = "FORUM"
			NewColumns5(10, DataType_Access)  = "TEXT (14)"
			NewColumns5(10, DataType_SQL6)  = "VARCHAR (14)"
			NewColumns5(10, DataType_SQL7)  = "NVARCHAR (14)"
			NewColumns5(10, DataType_MYSQL)  = "VARCHAR (14)"
			NewColumns5(10, ConstraintAccess)  = "NOT NULL"
			NewColumns5(10, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1027 DEFAULT ''"
			NewColumns5(10, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1027 DEFAULT ''"
			NewColumns5(10, ConstraintMySQL)  = "DEFAULT '' NOT NULL"

			NewColumns5(11, Prefix)	 = strTablePrefix
			NewColumns5(11, FieldName) = "F_DELETE_SCHED"
			NewColumns5(11, TableName) = "FORUM"
			NewColumns5(11, DataType_Access) = "INT"
			NewColumns5(11, DataType_SQL6) = "INT"
			NewColumns5(11, DataType_SQL7) = "INT"
			NewColumns5(11, DataType_MySQL) = "INT"
			NewColumns5(11, ConstraintAccess)  = "NOT NULL"
			NewColumns5(11, ConstraintSQL6)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1028 DEFAULT 365"
			NewColumns5(11, ConstraintSQL7)  = "NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1028 DEFAULT 365"
			NewColumns5(11, ConstraintMySQL)  = "DEFAULT '365' NOT NULL"

			call AddColumns(NewColumns5, intCriticalErrors, intWarnings)

			'## for Access we need to update the existing records !

			if strDBType = "access" then UpDateAccessFields(OldVersion)

			'## now transfer the config info from CONFIG to CONFIG_NEW

			TransferOldConfig

		end if

'###########################################################################
'##
'## Setup for update 6 / to version 3.3.03
'##
'###########################################################################

		if (OldVersion <= 6) then

  			Dim NewColumns6(3,11)

			NewColumns6(0, Prefix)	 = strTablePrefix
			NewColumns6(0, FieldName) = "F_A_COUNT"
			NewColumns6(0, TableName) = "FORUM"
			NewColumns6(0, DataType_Access) = "INT"
			NewColumns6(0, DataType_SQL6) = "INT"
			NewColumns6(0, DataType_SQL7) = "INT"
			NewColumns6(0, DataType_MySQL) = "INT"
			NewColumns6(0, ConstraintAccess)  = " NULL"
			NewColumns6(0, ConstraintSQL6)  = " NULL"
			NewColumns6(0, ConstraintSQL7)  = " NULL"
			NewColumns6(0, ConstraintMySQL)  = " NULL"

			NewColumns6(1, Prefix)	 = strTablePrefix
			NewColumns6(1, FieldName) = "F_A_TOPICS"
			NewColumns6(1, TableName) = "FORUM"
			NewColumns6(1, DataType_Access) = "INT"
			NewColumns6(1, DataType_SQL6) = "INT"
			NewColumns6(1, DataType_SQL7) = "INT"
			NewColumns6(1, DataType_MySQL) = "INT"
			NewColumns6(1, ConstraintAccess)  = " NULL"
			NewColumns6(1, ConstraintSQL6)  = " NULL"
			NewColumns6(1, ConstraintSQL7)  = " NULL"
			NewColumns6(1, ConstraintMySQL)  = " NULL"

			NewColumns6(2, Prefix)	 = strTablePrefix
			NewColumns6(2, FieldName) = "T_A_COUNT"
			NewColumns6(2, TableName) = "TOTALS"
			NewColumns6(2, DataType_Access) = "INT"
			NewColumns6(2, DataType_SQL6) = "INT"
			NewColumns6(2, DataType_SQL7) = "INT"
			NewColumns6(2, DataType_MySQL) = "INT"
			NewColumns6(2, ConstraintAccess)  = " NULL"
			NewColumns6(2, ConstraintSQL6)  = " NULL"
			NewColumns6(2, ConstraintSQL7)  = " NULL"
			NewColumns6(2, ConstraintMySQL)  = " NULL"

			NewColumns6(3, Prefix)	 = strTablePrefix
			NewColumns6(3, FieldName) = "P_A_COUNT"
			NewColumns6(3, TableName) = "TOTALS"
			NewColumns6(3, DataType_Access) = "INT"
			NewColumns6(3, DataType_SQL6) = "INT"
			NewColumns6(3, DataType_SQL7) = "INT"
			NewColumns6(3, DataType_MySQL) = "INT"
			NewColumns6(3, ConstraintAccess)  = " NULL"
			NewColumns6(3, ConstraintSQL6)  = " NULL"
			NewColumns6(3, ConstraintSQL7)  = " NULL"
			NewColumns6(3, ConstraintMySQL)  = " NULL"

			call AddColumns(NewColumns6, intCriticalErrors, intWarnings)

			'## Drop FORUM_A_TOPICS and recreate if needed
			'## Drop FORUM_A_REPLY and recreate if needed

			Dim SpecialSql6(4)

			strSql = "SELECT * FROM " & strTablePrefix & "A_TOPICS"

			set rs = my_Conn.Execute(strSql)

			if rs.eof then
				rs.close

				my_Conn.Execute("DROP TABLE " & strTablePrefix & "A_TOPICS")

				SpecialSql6(Access) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 				SpecialSql6(Access) = SpecialSql6(Access) & "CAT_ID int NOT NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "FORUM_ID int NOT NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "TOPIC_ID int NOT NULL , "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "T_STATUS smallint NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_MAIL smallint NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_SUBJECT varchar (100) NULL , "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "T_MESSAGE text NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_AUTHOR int NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_REPLIES int NULL , "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "T_VIEW_COUNT int NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_LAST_POST varchar (14) NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_DATE varchar (14) NULL, "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "T_LAST_POSTER int NULL, "
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_IP varchar (15) NULL, " 
 				SpecialSql6(Access) = SpecialSql6(Access) & "T_LAST_POST_AUTHOR int NULL ) "

				SpecialSql6(SQL6) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "CAT_ID int NOT NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "FORUM_ID int NOT NULL , "
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "TOPIC_ID int NOT NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_STATUS smallint NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_MAIL smallint NULL , "
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_SUBJECT varchar (100) NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_MESSAGE text NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_AUTHOR int NULL , "
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_REPLIES int NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_VIEW_COUNT int NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_LAST_POST varchar (14) NULL , "
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_DATE varchar (14) NULL, "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_LAST_POSTER int NULL, "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_IP varchar (15) NULL, " 
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "T_LAST_POST_AUTHOR int NULL ) "

				SpecialSql6(SQL7) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "CAT_ID int NOT NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "FORUM_ID int NOT NULL , "
	 			SpecialSql6(SQL7) = SpecialSql6(SQL7) & "TOPIC_ID int NOT NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_STATUS smallint NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_MAIL smallint NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_SUBJECT nvarchar (100) NULL , "
	 			SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_MESSAGE ntext NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_AUTHOR int NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_REPLIES int NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_VIEW_COUNT int NULL , "
	 			SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_LAST_POST varchar (14) NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_DATE varchar (14) NULL, "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_LAST_POSTER int NULL, "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_IP varchar (15) NULL, " 
	 			SpecialSql6(SQL7) = SpecialSql6(SQL7) & "T_LAST_POST_AUTHOR int NULL ) "

		 		SpecialSql6(MySql) = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
 				SpecialSql6(MySql) = SpecialSql6(MySql) & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
 				SpecialSql6(MySql) = SpecialSql6(MySql) & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "TOPIC_ID int (11) DEFAULT '' NOT NULL, "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_STATUS smallint (6) DEFAULT '1' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_MAIL smallint (6) DEFAULT '0' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_SUBJECT VARCHAR (100) DEFAULT '' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_MESSAGE text , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_AUTHOR int (11) DEFAULT '1' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_REPLIES int (11) DEFAULT '0' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_VIEW_COUNT int (11) DEFAULT '0' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_LAST_POST VARCHAR (14) DEFAULT '' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_DATE VARCHAR (14) DEFAULT '', "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_LAST_POSTER int (11) DEFAULT '1', "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_IP VARCHAR (15) DEFAULT '000.000.000.000', " 
				SpecialSql6(MySql) = SpecialSql6(MySql) & "T_LAST_POST_AUTHOR int (11) DEFAULT '1',   "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_TOPIC_CATFORTOP(CAT_ID,FORUM_ID,TOPIC_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_TOPIC_CAT_ID(CAT_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_TOPIC_FORUM_ID(FORUM_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_TOPIC_TOPIC_ID (TOPIC_ID) )"

		 		strOkMessage = "Table A_TOPICS re-created "

				call SpecialUpdates(SpecialSql6, strOkMessage)

				Response.Flush
			else
				rs.close
			end if

			strSql = "SELECT * FROM " & strTablePrefix & "A_REPLY"

			set rs = my_Conn.Execute(strSql)

			if rs.eof then
				rs.close
				my_Conn.Execute("DROP TABLE " & strTablePrefix & "A_REPLY")

		 		SpecialSql6(Access) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
 				SpecialSql6(Access) = SpecialSql6(Access) & "CAT_ID int NOT NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "FORUM_ID int NOT NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "TOPIC_ID int NOT NULL , "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "REPLY_ID int NOT NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "R_STATUS smallint NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "R_MAIL smallint NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "R_AUTHOR int NULL , "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "R_MESSAGE text NULL , "
 				SpecialSql6(Access) = SpecialSql6(Access) & "R_DATE varchar (14) NULL , "
	 			SpecialSql6(Access) = SpecialSql6(Access) & "R_IP varchar (15) NULL ) "

	 			SpecialSql6(SQL6) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "CAT_ID int NOT NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "FORUM_ID int NOT NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "TOPIC_ID int NOT NULL , "
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "REPLY_ID int NOT NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "R_MAIL smallint NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "R_STATUS smallint NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "R_AUTHOR int NULL , "
	 			SpecialSql6(SQL6) = SpecialSql6(SQL6) & "R_MESSAGE text NULL , "
				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "R_DATE varchar (14) NULL , "
 				SpecialSql6(SQL6) = SpecialSql6(SQL6) & "R_IP varchar (15) NULL ) "

		 		SpecialSql6(SQL7) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "CAT_ID int NOT NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "FORUM_ID int NOT NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "TOPIC_ID int NOT NULL , "
	 			SpecialSql6(SQL7) = SpecialSql6(SQL7) & "REPLY_ID int NOT NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "R_STATUS smallint NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "R_MAIL smallint NULL , "
	 			SpecialSql6(SQL7) = SpecialSql6(SQL7) & "R_AUTHOR int NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "R_MESSAGE ntext NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "R_DATE varchar (14) NULL , "
 				SpecialSql6(SQL7) = SpecialSql6(SQL7) & "R_IP varchar (15) NULL ) "

				SpecialSql6(MySql) = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "TOPIC_ID int (11) DEFAULT '1' NOT NULL , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "REPLY_ID int (11) DEFAULT '' NOT NULL , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "R_STATUS smallint (6) DEFAULT '1' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "R_AUTHOR int (11) DEFAULT '1' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "R_MESSAGE text , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "R_DATE VARCHAR (14) DEFAULT '' , "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "R_IP VARCHAR (15) DEFAULT '000.000.000.000', "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_REPLY_CATFORTOPREPL(CAT_ID,FORUM_ID,TOPIC_ID, REPLY_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_REPLY_REP_ID(REPLY_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_REPLY_CAT_ID(CAT_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_REPLY_FORUM_ID(FORUM_ID), "
				SpecialSql6(MySql) = SpecialSql6(MySql) & "KEY " & strTablePrefix & "A_REPLY_TOPIC_ID (TOPIC_ID) )"	

		 		strOkMessage = "Table A_REPLY re-created "

				call SpecialUpdates(SpecialSql6, strOkMessage)

				Response.Flush
			else
				rs.close
				'## Add the missing R_STATUS field to the Access database
				if strDBType = "access" then 
			 		strSql = "ALTER TABLE " & strTablePrefix & "A_REPLY "
	 				strSql = strSql & "ADD R_STATUS smallint NULL "	

					on error resume next
					my_Conn.Errors.Clear
					Err.Clear
					my_Conn.Execute (strSql)

					Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

					UpdateErrorCode = UpdateErrorCheck()

					on error goto 0

					if UpdateErrorCode = 0 then
						Response.Write("  <tr>" & vbNewLine)
						Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
						Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> R_STATUS field added to " & strTablePrefix & "A_REPLY</font></td>" & vbNewLine)
						Response.Write("  </tr>" & vbNewLine)
					elseif UpdateErrorCode = 1 then
						Response.Write("  <tr>" & vbNewLine)
						Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Noncritical error: </b></font></td>" & vbNewLine)
						Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> R_STATUS already existed in " & strTablePrefix & "A_REPLY</font></td>" & vbNewLine)
						Response.Write("  </tr>" & vbNewLine)
						intWarnings = intWarnings + 1
					elseif UpdateErrorCode = 2 then
						Response.Write("  <tr>" & vbNewLine)
						Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
						Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> No write access to the table " & strTablePrefix & "A_REPLY<br />R_STATUS not added to database!</font></td>" & vbNewLine)
						Response.Write("  </tr>" & vbNewLine)
						intCriticalErrors = intCriticalErrors + 1
					else
						Response.Write("  <tr>" & vbNewLine)
						Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
						Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " in statement [" & strUpdateSql & "] while trying to add R_STATUS to " & strTablePrefix & "A_REPLY</font></td>" & vbNewLine)
						Response.Write("  </tr>" & vbNewLine)
						intCriticalErrors = intCriticalErrors + 1
					end if
					Response.Write("</table>" & vbNewLine)
				end if
			end if
			set rs = nothing

			'## Add the missing config-values to the database if needed

			strDummy = SetConfigValue(0,"STRSUBSCRIPTION", "1")
			strDummy = SetConfigValue(0,"STRMODERATION", "1")

			'## update the status of archived replies...

			strSql = "UPDATE " & strTablePrefix & "A_REPLY "
			strSql = strSql & " SET R_STATUS = 0 "
			strSql = strSql & " WHERE (R_STATUS IS NULL)"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Status of archived replies updated</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to update the status of the archived replies</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)

			'## update the version info...	

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if

'###########################################################################
'## 
'## Setup for update 7 / to version 3.3.04
'## 
'###########################################################################

		if (OldVersion <= 7) then

			if strDBType = "access" then

				Dim SpecialSql7(1)

				'## Change T_MESSAGE to a MEMO/TEXT field in A_TOPICS table

		 		SpecialSql7(Access) = "ALTER TABLE " & strTablePrefix & "A_TOPICS "
				SpecialSql7(Access) = SpecialSql7(Access) & "ALTER COLUMN T_MESSAGE MEMO NULL "

 				strOkMessage = "T_MESSAGE Field has been changed"

				call SpecialUpdates(SpecialSql7, strOkMessage)

				Response.Flush

				'## Change R_MESSAGE to a MEMO/TEXT field in A_REPLY table

	 			SpecialSql7(Access) = "ALTER TABLE " & strTablePrefix & "A_REPLY "
				SpecialSql7(Access) = SpecialSql7(Access) & "ALTER COLUMN R_MESSAGE MEMO NULL "

 				strOkMessage = "R_MESSAGE Field has been changed"

				call SpecialUpdates(SpecialSql7, strOkMessage)

				Response.Flush
			end if

			'## update the version info...	

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
			elseif UpdateErrorCode = 2 then
				Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td></tr>")
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td></tr>")
				intCriticalErrors = intCriticalErrors + 1
			end if

			Response.Write("</table>")
			Response.Flush
		end if

'###########################################################################
'## 
'## Setup for update 7 / to version 3.3.05 (just a version # change)
'## 
'###########################################################################

		if (OldVersion <= 7) then

			'## update the version info...	

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
			elseif UpdateErrorCode = 2 then
				Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td></tr>")
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td></tr>")
				intCriticalErrors = intCriticalErrors + 1
			end if

			Response.Write("</table>")
			Response.Flush
		end if

'###########################################################################
'##
'## Setup for update 8 / to version 3.4
'##
'###########################################################################

		if (OldVersion <= 8) then
  			Dim NewColumns8(29,11)

			NewColumns8(0, Prefix)	 = strMemberTablePrefix
			NewColumns8(0, FieldName) = "M_MSN"
			NewColumns8(0, TableName) = "MEMBERS"
			NewColumns8(0, DataType_Access) = "TEXT (150)"
			NewColumns8(0, DataType_SQL6) = "VARCHAR (150)"
			NewColumns8(0, DataType_SQL7) = "NVARCHAR (150)"
			NewColumns8(0, DataType_MySQL) = "VARCHAR (150)"
			NewColumns8(0, ConstraintAccess)  = " NULL"
			NewColumns8(0, ConstraintSQL6)  = " NULL DEFAULT ''"
			NewColumns8(0, ConstraintSQL7)  = " NULL DEFAULT ''"
			NewColumns8(0, ConstraintMySQL)  = " DEFAULT '' NULL"

			NewColumns8(1, Prefix)	 = strMemberTablePrefix
			NewColumns8(1, FieldName) = "M_KEY"
			NewColumns8(1, TableName) = "MEMBERS"
			NewColumns8(1, DataType_Access) = "TEXT (32)"
			NewColumns8(1, DataType_SQL6) = "VARCHAR (32)"
			NewColumns8(1, DataType_SQL7) = "NVARCHAR (32)"
			NewColumns8(1, DataType_MySQL) = "VARCHAR (32)"
			NewColumns8(1, ConstraintAccess)  = " NULL"
			NewColumns8(1, ConstraintSQL6)  = " NULL DEFAULT ''"
			NewColumns8(1, ConstraintSQL7)  = " NULL DEFAULT ''"
			NewColumns8(1, ConstraintMySQL)  = " DEFAULT '' NULL"

			NewColumns8(2, Prefix)	 = strMemberTablePrefix
			NewColumns8(2, FieldName) = "M_NEWEMAIL"
			NewColumns8(2, TableName) = "MEMBERS"
			NewColumns8(2, DataType_Access) = "TEXT (50)"
			NewColumns8(2, DataType_SQL6) = "VARCHAR (50)"
			NewColumns8(2, DataType_SQL7) = "NVARCHAR (50)"
			NewColumns8(2, DataType_MySQL) = "VARCHAR (50)"
			NewColumns8(2, ConstraintAccess)  = " NULL"
			NewColumns8(2, ConstraintSQL6)  = " NULL DEFAULT ''"
			NewColumns8(2, ConstraintSQL7)  = " NULL DEFAULT ''"
			NewColumns8(2, ConstraintMySQL)  = " DEFAULT '' NULL"

			NewColumns8(3, Prefix)	 = strMemberTablePrefix
			NewColumns8(3, FieldName) = "M_SHA256"
			NewColumns8(3, TableName) = "MEMBERS"
			NewColumns8(3, DataType_Access) = "smallint"
			NewColumns8(3, DataType_SQL6) = "smallint"
			NewColumns8(3, DataType_SQL7) = "smallint"
			NewColumns8(3, DataType_MySQL) = "smallint (6)"
			NewColumns8(3, ConstraintAccess)  = " NULL"
			NewColumns8(3, ConstraintSQL6)  = " NULL"
			NewColumns8(3, ConstraintSQL7)  = " NULL"
			NewColumns8(3, ConstraintMySQL)  = " NULL"

			NewColumns8(4, Prefix)	 = strMemberTablePrefix
			NewColumns8(4, FieldName) = "M_PWKEY"
			NewColumns8(4, TableName) = "MEMBERS"
			NewColumns8(4, DataType_Access) = "TEXT (32)"
			NewColumns8(4, DataType_SQL6) = "VARCHAR (32)"
			NewColumns8(4, DataType_SQL7) = "NVARCHAR (32)"
			NewColumns8(4, DataType_MySQL) = "VARCHAR (32)"
			NewColumns8(4, ConstraintAccess)  = " NULL"
			NewColumns8(4, ConstraintSQL6)  = " NULL DEFAULT ''"
			NewColumns8(4, ConstraintSQL7)  = " NULL DEFAULT ''"
			NewColumns8(4, ConstraintMySQL)  = " DEFAULT '' NULL"

			NewColumns8(5, Prefix)	 = strTablePrefix
			NewColumns8(5, FieldName) = "T_STICKY"
			NewColumns8(5, TableName) = "TOPICS"
			NewColumns8(5, DataType_Access) = "smallint"
			NewColumns8(5, DataType_SQL6) = "smallint"
			NewColumns8(5, DataType_SQL7) = "smallint"
			NewColumns8(5, DataType_MySQL) = "smallint (6)"
			NewColumns8(5, ConstraintAccess)  = " NULL"
			NewColumns8(5, ConstraintSQL6)  = " NULL DEFAULT 0"
			NewColumns8(5, ConstraintSQL7)  = " NULL DEFAULT 0"
			NewColumns8(5, ConstraintMySQL)  = " DEFAULT '0' NULL"

			NewColumns8(6, Prefix)	 = strTablePrefix
			NewColumns8(6, FieldName) = "T_STICKY"
			NewColumns8(6, TableName) = "A_TOPICS"
			NewColumns8(6, DataType_Access) = "smallint"
			NewColumns8(6, DataType_SQL6) = "smallint"
			NewColumns8(6, DataType_SQL7) = "smallint"
			NewColumns8(6, DataType_MySQL) = "smallint (6)"
			NewColumns8(6, ConstraintAccess)  = " NULL"
			NewColumns8(6, ConstraintSQL6)  = " NULL"
			NewColumns8(6, ConstraintSQL7)  = " NULL"
			NewColumns8(6, ConstraintMySQL)  = " NULL"

			NewColumns8(7, Prefix)	 = strTablePrefix
			NewColumns8(7, FieldName) = "T_LAST_EDIT"
			NewColumns8(7, TableName) = "TOPICS"
			NewColumns8(7, DataType_Access) = "TEXT (14)"
			NewColumns8(7, DataType_SQL6) = "VARCHAR (14)"
			NewColumns8(7, DataType_SQL7) = "NVARCHAR (14)"
			NewColumns8(7, DataType_MySQL) = "VARCHAR (14)"
			NewColumns8(7, ConstraintAccess)  = " NULL"
			NewColumns8(7, ConstraintSQL6)  = " NULL"
			NewColumns8(7, ConstraintSQL7)  = " NULL"
			NewColumns8(7, ConstraintMySQL)  = " NULL"

			NewColumns8(8, Prefix)	 = strTablePrefix
			NewColumns8(8, FieldName) = "T_LAST_EDIT"
			NewColumns8(8, TableName) = "A_TOPICS"
			NewColumns8(8, DataType_Access) = "TEXT (14)"
			NewColumns8(8, DataType_SQL6) = "VARCHAR (14)"
			NewColumns8(8, DataType_SQL7) = "NVARCHAR (14)"
			NewColumns8(8, DataType_MySQL) = "VARCHAR (14)"
			NewColumns8(8, ConstraintAccess)  = " NULL"
			NewColumns8(8, ConstraintSQL6)  = " NULL"
			NewColumns8(8, ConstraintSQL7)  = " NULL"
			NewColumns8(8, ConstraintMySQL)  = " NULL"

			NewColumns8(9, Prefix)	 = strTablePrefix
			NewColumns8(9, FieldName) = "T_LAST_EDITBY"
			NewColumns8(9, TableName) = "TOPICS"
			NewColumns8(9, DataType_Access) = "INT"
			NewColumns8(9, DataType_SQL6) = "INT"
			NewColumns8(9, DataType_SQL7) = "INT"
			NewColumns8(9, DataType_MySQL) = "INT (11)"
			NewColumns8(9, ConstraintAccess)  = " NULL"
			NewColumns8(9, ConstraintSQL6)  = " NULL"
			NewColumns8(9, ConstraintSQL7)  = " NULL"
			NewColumns8(9, ConstraintMySQL)  = " NULL"

			NewColumns8(10, Prefix)	 = strTablePrefix
			NewColumns8(10, FieldName) = "T_LAST_EDITBY"
			NewColumns8(10, TableName) = "A_TOPICS"
			NewColumns8(10, DataType_Access) = "INT"
			NewColumns8(10, DataType_SQL6) = "INT"
			NewColumns8(10, DataType_SQL7) = "INT"
			NewColumns8(10, DataType_MySQL) = "INT (11)"
			NewColumns8(10, ConstraintAccess)  = " NULL"
			NewColumns8(10, ConstraintSQL6)  = " NULL"
			NewColumns8(10, ConstraintSQL7)  = " NULL"
			NewColumns8(10, ConstraintMySQL)  = " NULL"

			NewColumns8(11, Prefix)	 = strTablePrefix
			NewColumns8(11, FieldName) = "R_LAST_EDIT"
			NewColumns8(11, TableName) = "REPLY"
			NewColumns8(11, DataType_Access) = "TEXT (14)"
			NewColumns8(11, DataType_SQL6) = "VARCHAR (14)"
			NewColumns8(11, DataType_SQL7) = "NVARCHAR (14)"
			NewColumns8(11, DataType_MySQL) = "VARCHAR (14)"
			NewColumns8(11, ConstraintAccess)  = " NULL"
			NewColumns8(11, ConstraintSQL6)  = " NULL"
			NewColumns8(11, ConstraintSQL7)  = " NULL"
			NewColumns8(11, ConstraintMySQL)  = " NULL"

			NewColumns8(12, Prefix)	 = strTablePrefix
			NewColumns8(12, FieldName) = "R_LAST_EDIT"
			NewColumns8(12, TableName) = "A_REPLY"
			NewColumns8(12, DataType_Access) = "TEXT (14)"
			NewColumns8(12, DataType_SQL6) = "VARCHAR (14)"
			NewColumns8(12, DataType_SQL7) = "NVARCHAR (14)"
			NewColumns8(12, DataType_MySQL) = "VARCHAR (14)"
			NewColumns8(12, ConstraintAccess)  = " NULL"
			NewColumns8(12, ConstraintSQL6)  = " NULL"
			NewColumns8(12, ConstraintSQL7)  = " NULL"
			NewColumns8(12, ConstraintMySQL)  = " NULL"

			NewColumns8(13, Prefix)	 = strTablePrefix
			NewColumns8(13, FieldName) = "R_LAST_EDITBY"
			NewColumns8(13, TableName) = "REPLY"
			NewColumns8(13, DataType_Access) = "INT"
			NewColumns8(13, DataType_SQL6) = "INT"
			NewColumns8(13, DataType_SQL7) = "INT"
			NewColumns8(13, DataType_MySQL) = "INT (11)"
			NewColumns8(13, ConstraintAccess)  = " NULL"
			NewColumns8(13, ConstraintSQL6)  = " NULL"
			NewColumns8(13, ConstraintSQL7)  = " NULL"
			NewColumns8(13, ConstraintMySQL)  = " NULL"

			NewColumns8(14, Prefix)	 = strTablePrefix
			NewColumns8(14, FieldName) = "R_LAST_EDITBY"
			NewColumns8(14, TableName) = "A_REPLY"
			NewColumns8(14, DataType_Access) = "INT"
			NewColumns8(14, DataType_SQL6) = "INT"
			NewColumns8(14, DataType_SQL7) = "INT"
			NewColumns8(14, DataType_MySQL) = "INT (11)"
			NewColumns8(14, ConstraintAccess)  = " NULL"
			NewColumns8(14, ConstraintSQL6)  = " NULL"
			NewColumns8(14, ConstraintSQL7)  = " NULL"
			NewColumns8(14, ConstraintMySQL)  = " NULL"

			NewColumns8(15, Prefix)	 = strTablePrefix
			NewColumns8(15, FieldName) = "T_SIG"
			NewColumns8(15, TableName) = "TOPICS"
			NewColumns8(15, DataType_Access) = "smallint"
			NewColumns8(15, DataType_SQL6) = "smallint"
			NewColumns8(15, DataType_SQL7) = "smallint"
			NewColumns8(15, DataType_MySQL) = "smallint (6)"
			NewColumns8(15, ConstraintAccess)  = " NULL"
			NewColumns8(15, ConstraintSQL6)  = " NULL DEFAULT 0"
			NewColumns8(15, ConstraintSQL7)  = " NULL DEFAULT 0"
			NewColumns8(15, ConstraintMySQL)  = " DEFAULT '0' NULL"

			NewColumns8(16, Prefix)	 = strTablePrefix
			NewColumns8(16, FieldName) = "T_SIG"
			NewColumns8(16, TableName) = "A_TOPICS"
			NewColumns8(16, DataType_Access) = "smallint"
			NewColumns8(16, DataType_SQL6) = "smallint"
			NewColumns8(16, DataType_SQL7) = "smallint"
			NewColumns8(16, DataType_MySQL) = "smallint (6)"
			NewColumns8(16, ConstraintAccess)  = " NULL"
			NewColumns8(16, ConstraintSQL6)  = " NULL"
			NewColumns8(16, ConstraintSQL7)  = " NULL"
			NewColumns8(16, ConstraintMySQL)  = " NULL"

			NewColumns8(17, Prefix)	 = strTablePrefix
			NewColumns8(17, FieldName) = "R_SIG"
			NewColumns8(17, TableName) = "REPLY"
			NewColumns8(17, DataType_Access) = "smallint"
			NewColumns8(17, DataType_SQL6) = "smallint"
			NewColumns8(17, DataType_SQL7) = "smallint"
			NewColumns8(17, DataType_MySQL) = "smallint (6)"
			NewColumns8(17, ConstraintAccess)  = " NULL"
			NewColumns8(17, ConstraintSQL6)  = " NULL DEFAULT 0"
			NewColumns8(17, ConstraintSQL7)  = " NULL DEFAULT 0"
			NewColumns8(17, ConstraintMySQL)  = " DEFAULT '0' NULL"

			NewColumns8(18, Prefix)	 = strTablePrefix
			NewColumns8(18, FieldName) = "R_SIG"
			NewColumns8(18, TableName) = "A_REPLY"
			NewColumns8(18, DataType_Access) = "smallint"
			NewColumns8(18, DataType_SQL6) = "smallint"
			NewColumns8(18, DataType_SQL7) = "smallint"
			NewColumns8(18, DataType_MySQL) = "smallint (6)"
			NewColumns8(18, ConstraintAccess)  = " NULL"
			NewColumns8(18, ConstraintSQL6)  = " NULL"
			NewColumns8(18, ConstraintSQL7)  = " NULL"
			NewColumns8(18, ConstraintMySQL)  = " NULL"

			NewColumns8(19, Prefix)	 = strMemberTablePrefix
			NewColumns8(19, FieldName) = "M_VIEW_SIG"
			NewColumns8(19, TableName) = "MEMBERS"
			NewColumns8(19, DataType_Access) = "smallint"
			NewColumns8(19, DataType_SQL6) = "smallint"
			NewColumns8(19, DataType_SQL7) = "smallint"
			NewColumns8(19, DataType_MySQL) = "smallint (6)"
			NewColumns8(19, ConstraintAccess)  = " NULL"
			NewColumns8(19, ConstraintSQL6)  = " NULL DEFAULT 1"
			NewColumns8(19, ConstraintSQL7)  = " NULL DEFAULT 1"
			NewColumns8(19, ConstraintMySQL)  = " DEFAULT '1' NULL"

			NewColumns8(20, Prefix)	 = strTablePrefix
			NewColumns8(20, FieldName) = "F_DEFAULTDAYS"
			NewColumns8(20, TableName) = "FORUM"
			NewColumns8(20, DataType_Access) = "int"
			NewColumns8(20, DataType_SQL6) = "int"
			NewColumns8(20, DataType_SQL7) = "int"
			NewColumns8(20, DataType_MySQL) = "int (11)"
			NewColumns8(20, ConstraintAccess)  = " NULL"
			NewColumns8(20, ConstraintSQL6)  = " NULL DEFAULT 30"
			NewColumns8(20, ConstraintSQL7)  = " NULL DEFAULT 30"
			NewColumns8(20, ConstraintMySQL)  = " DEFAULT '30' NULL"

			NewColumns8(21, Prefix)	 = strTablePrefix
			NewColumns8(21, FieldName) = "F_COUNT_M_POSTS"
			NewColumns8(21, TableName) = "FORUM"
			NewColumns8(21, DataType_Access) = "smallint"
			NewColumns8(21, DataType_SQL6) = "smallint"
			NewColumns8(21, DataType_SQL7) = "smallint"
			NewColumns8(21, DataType_MySQL) = "smallint (6)"
			NewColumns8(21, ConstraintAccess)  = " NULL"
			NewColumns8(21, ConstraintSQL6)  = " NULL DEFAULT 1"
			NewColumns8(21, ConstraintSQL7)  = " NULL DEFAULT 1"
			NewColumns8(21, ConstraintMySQL)  = " DEFAULT '1' NULL"

			NewColumns8(22, Prefix)	 = strMemberTablePrefix
			NewColumns8(22, FieldName) = "M_DOB"
			NewColumns8(22, TableName) = "MEMBERS"
			NewColumns8(22, DataType_Access) = "TEXT (8)"
			NewColumns8(22, DataType_SQL6) = "VARCHAR (8)"
			NewColumns8(22, DataType_SQL7) = "NVARCHAR (8)"
			NewColumns8(22, DataType_MySQL) = "VARCHAR (8)"
			NewColumns8(22, ConstraintAccess)  = " NULL"
			NewColumns8(22, ConstraintSQL6)  = " NULL DEFAULT ''"
			NewColumns8(22, ConstraintSQL7)  = " NULL DEFAULT ''"
			NewColumns8(22, ConstraintMySQL)  = " DEFAULT '' NULL"

			NewColumns8(23, Prefix)	 = strTablePrefix
			NewColumns8(23, FieldName) = "F_LAST_POST_TOPIC_ID"
			NewColumns8(23, TableName) = "FORUM"
			NewColumns8(23, DataType_Access) = "int"
			NewColumns8(23, DataType_SQL6) = "int"
			NewColumns8(23, DataType_SQL7) = "int"
			NewColumns8(23, DataType_MySQL) = "int (11)"
			NewColumns8(23, ConstraintAccess)  = " NULL"
			NewColumns8(23, ConstraintSQL6)  = " NULL DEFAULT 0"
			NewColumns8(23, ConstraintSQL7)  = " NULL DEFAULT 0"
			NewColumns8(23, ConstraintMySQL)  = " DEFAULT '0' NULL"

			NewColumns8(24, Prefix)	 = strTablePrefix
			NewColumns8(24, FieldName) = "F_LAST_POST_REPLY_ID"
			NewColumns8(24, TableName) = "FORUM"
			NewColumns8(24, DataType_Access) = "int"
			NewColumns8(24, DataType_SQL6) = "int"
			NewColumns8(24, DataType_SQL7) = "int"
			NewColumns8(24, DataType_MySQL) = "int (11)"
			NewColumns8(24, ConstraintAccess)  = " NULL"
			NewColumns8(24, ConstraintSQL6)  = " NULL DEFAULT 0"
			NewColumns8(24, ConstraintSQL7)  = " NULL DEFAULT 0"
			NewColumns8(24, ConstraintMySQL)  = " DEFAULT '0' NULL"

			NewColumns8(25, Prefix)	 = strTablePrefix
			NewColumns8(25, FieldName) = "T_LAST_POST_REPLY_ID"
			NewColumns8(25, TableName) = "TOPICS"
			NewColumns8(25, DataType_Access) = "int"
			NewColumns8(25, DataType_SQL6) = "int"
			NewColumns8(25, DataType_SQL7) = "int"
			NewColumns8(25, DataType_MySQL) = "int (11)"
			NewColumns8(25, ConstraintAccess)  = " NULL"
			NewColumns8(25, ConstraintSQL6)  = " NULL DEFAULT 0"
			NewColumns8(25, ConstraintSQL7)  = " NULL DEFAULT 0"
			NewColumns8(25, ConstraintMySQL)  = " DEFAULT '0' NULL"

			NewColumns8(26, Prefix)	 = strTablePrefix
			NewColumns8(26, FieldName) = "T_LAST_POST_REPLY_ID"
			NewColumns8(26, TableName) = "A_TOPICS"
			NewColumns8(26, DataType_Access) = "int"
			NewColumns8(26, DataType_SQL6) = "int"
			NewColumns8(26, DataType_SQL7) = "int"
			NewColumns8(26, DataType_MySQL) = "int (11)"
			NewColumns8(26, ConstraintAccess)  = " NULL"
			NewColumns8(26, ConstraintSQL6)  = " NULL"
			NewColumns8(26, ConstraintSQL7)  = " NULL"
			NewColumns8(26, ConstraintMySQL)  = " NULL"

			NewColumns8(27, Prefix)	 = strTablePrefix
			NewColumns8(27, FieldName) = "T_UREPLIES"
			NewColumns8(27, TableName) = "TOPICS"
			NewColumns8(27, DataType_Access) = "int"
			NewColumns8(27, DataType_SQL6) = "int"
			NewColumns8(27, DataType_SQL7) = "int"
			NewColumns8(27, DataType_MySQL) = "int (11)"
			NewColumns8(27, ConstraintAccess)  = " NULL"
			NewColumns8(27, ConstraintSQL6)  = " NULL"
			NewColumns8(27, ConstraintSQL7)  = " NULL"
			NewColumns8(27, ConstraintMySQL)  = " NULL"

			NewColumns8(28, Prefix)	 = strTablePrefix
			NewColumns8(28, FieldName) = "T_UREPLIES"
			NewColumns8(28, TableName) = "A_TOPICS"
			NewColumns8(28, DataType_Access) = "int"
			NewColumns8(28, DataType_SQL6) = "int"
			NewColumns8(28, DataType_SQL7) = "int"
			NewColumns8(28, DataType_MySQL) = "int (11)"
			NewColumns8(28, ConstraintAccess)  = " NULL"
			NewColumns8(28, ConstraintSQL6)  = " NULL"
			NewColumns8(28, ConstraintSQL7)  = " NULL"
			NewColumns8(28, ConstraintMySQL)  = " NULL"

			NewColumns8(29, Prefix)	 = strMemberTablePrefix
			NewColumns8(29, FieldName) = "M_SIG_DEFAULT"
			NewColumns8(29, TableName) = "MEMBERS"
			NewColumns8(29, DataType_Access) = "smallint"
			NewColumns8(29, DataType_SQL6) = "smallint"
			NewColumns8(29, DataType_SQL7) = "smallint"
			NewColumns8(29, DataType_MySQL) = "smallint (6)"
			NewColumns8(29, ConstraintAccess)  = " NULL"
			NewColumns8(29, ConstraintSQL6)  = " NULL DEFAULT 1"
			NewColumns8(29, ConstraintSQL7)  = " NULL DEFAULT 1"
			NewColumns8(29, ConstraintMySQL)  = " DEFAULT '1' NULL"

			call AddColumns(NewColumns8, intCriticalErrors, intWarnings)

			Dim SpecialSql8(4)

			'## Update F_DEFAULTDAYS for existing Forums in FORUM table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "FORUM SET F_DEFAULTDAYS = 30 WHERE (F_DEFAULTDAYS IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "FORUM SET F_DEFAULTDAYS = 30 WHERE (F_DEFAULTDAYS IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "FORUM SET F_DEFAULTDAYS = 30 WHERE (F_DEFAULTDAYS IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "FORUM SET F_DEFAULTDAYS = 30 WHERE (F_DEFAULTDAYS IS NULL)"

			strOkMessage = "F_DEFAULTDAYS field value updated in the FORUM table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update F_COUNT_M_POSTS for existing Forums in FORUM table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "FORUM SET F_COUNT_M_POSTS = 1 WHERE (F_COUNT_M_POSTS IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "FORUM SET F_COUNT_M_POSTS = 1 WHERE (F_COUNT_M_POSTS IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "FORUM SET F_COUNT_M_POSTS = 1 WHERE (F_COUNT_M_POSTS IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "FORUM SET F_COUNT_M_POSTS = 1 WHERE (F_COUNT_M_POSTS IS NULL)"

			strOkMessage = "F_COUNT_M_POSTS field value updated in the FORUM table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update T_STICKY for existing Topics in TOPICS table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"

			strOkMessage = "T_STICKY field value updated in the TOPICS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update T_STICKY for existing Topics in A_TOPICS table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"

			strOkMessage = "T_STICKY field value updated in the A_TOPICS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update T_LAST_POST_REPLY_ID for existing Topics in TOPICS table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "TOPICS SET T_LAST_POST_REPLY_ID = 0 WHERE (T_LAST_POST_REPLY_ID IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "TOPICS SET T_LAST_POST_REPLY_ID = 0 WHERE (T_LAST_POST_REPLY_ID IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "TOPICS SET T_LAST_POST_REPLY_ID = 0 WHERE (T_LAST_POST_REPLY_ID IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "TOPICS SET T_LAST_POST_REPLY_ID = 0 WHERE (T_LAST_POST_REPLY_ID IS NULL)"

			strOkMessage = "T_LAST_POST_REPLY_ID field value updated in the TOPICS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update T_SIG for existing Topics in TOPICS table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"

			strOkMessage = "T_SIG field value updated in the TOPICS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update T_SIG for existing Topics in A_TOPICS table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "A_TOPICS SET T_SIG = 0 WHERE (T_SIG IS NULL)"

			strOkMessage = "T_SIG field value updated in the A_TOPICS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update R_SIG for existing Replies in REPLY table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"

			strOkMessage = "R_SIG field value updated in the REPLY table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update R_SIG for existing Replies in A_REPLY table

			SpecialSql8(Access) = "UPDATE " & strTablePrefix & "A_REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strTablePrefix & "A_REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strTablePrefix & "A_REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strTablePrefix & "A_REPLY SET R_SIG = 0 WHERE (R_SIG IS NULL)"

			strOkMessage = "R_SIG field value updated in the A_REPLY table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update M_VIEW_SIG for existing Members in MEMBERS table

			SpecialSql8(Access) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_VIEW_SIG = 1 WHERE (M_VIEW_SIG IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_VIEW_SIG = 1 WHERE (M_VIEW_SIG IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_VIEW_SIG = 1 WHERE (M_VIEW_SIG IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_VIEW_SIG = 1 WHERE (M_VIEW_SIG IS NULL)"

			strOkMessage = "M_VIEW_SIG field value updated in the MEMBERS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Update M_SIG_DEFAULT for existing Members in MEMBERS table

			SpecialSql8(Access) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_SIG_DEFAULT = 1 WHERE (M_SIG_DEFAULT IS NULL)"
			SpecialSql8(SQL6) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_SIG_DEFAULT = 1 WHERE (M_SIG_DEFAULT IS NULL)"
			SpecialSql8(SQL7) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_SIG_DEFAULT = 1 WHERE (M_SIG_DEFAULT IS NULL)"
			SpecialSql8(MySql) = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_SIG_DEFAULT = 1 WHERE (M_SIG_DEFAULT IS NULL)"

			strOkMessage = "M_SIG_DEFAULT field value updated in the MEMBERS table"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change M_KEY field size to 32 characters

	 		SpecialSql8(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN M_KEY TEXT (32) NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN M_KEY VARCHAR (32) NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN M_KEY NVARCHAR (32) NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY M_KEY VARCHAR (32) NULL "

	 		strOkMessage = "M_KEY Field Size has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change M_PWKEY field size to 32 characters

	 		SpecialSql8(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN M_PWKEY TEXT (32) NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN M_PWKEY VARCHAR (32) NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN M_PWKEY NVARCHAR (32) NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY M_PWKEY VARCHAR (32) NULL "

	 		strOkMessage = "M_PWKEY Field Size has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change M_SIG to a MEMO/TEXT field

	 		SpecialSql8(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN M_SIG MEMO NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN M_SIG TEXT NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN M_SIG NTEXT NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY M_SIG TEXT NULL "

	 		strOkMessage = "M_SIG Field has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change M_COUNTRY field size to 50 characters

	 		SpecialSql8(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN M_COUNTRY TEXT (50) NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN M_COUNTRY VARCHAR (50) NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN M_COUNTRY NVARCHAR (50) NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY M_COUNTRY VARCHAR (50) NULL "

	 		strOkMessage = "M_COUNTRY Field Size has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change M_HOMEPAGE field size to 255 characters

	 		SpecialSql8(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN M_HOMEPAGE TEXT (255) NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN M_HOMEPAGE VARCHAR (255) NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN M_HOMEPAGE NVARCHAR (255) NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY M_HOMEPAGE VARCHAR (255) NULL "

	 		strOkMessage = "M_HOMEPAGE Field Size has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change M_PASSWORD field size to 65 characters

	 		SpecialSql8(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN M_PASSWORD TEXT (65) NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN M_PASSWORD VARCHAR (65) NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN M_PASSWORD NVARCHAR (65) NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY M_PASSWORD VARCHAR (65) NULL "

	 		strOkMessage = "M_PASSWORD Field Size has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Change F_DESCRIPTION to a MEMO/TEXT field

	 		SpecialSql8(Access) = "ALTER TABLE " & strTablePrefix & "FORUM "
			SpecialSql8(Access) = SpecialSql8(Access) & "ALTER COLUMN F_DESCRIPTION MEMO NULL "

 			SpecialSql8(SQL6) = "ALTER TABLE " & strTablePrefix & "FORUM "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "ALTER COLUMN F_DESCRIPTION TEXT NULL "

	 		SpecialSql8(SQL7) = "ALTER TABLE " & strTablePrefix & "FORUM "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "ALTER COLUMN F_DESCRIPTION NTEXT NULL "

			SpecialSql8(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MODIFY F_DESCRIPTION TEXT NULL "

	 		strOkMessage = "F_DESCRIPTION Field has been changed"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Drop Table MEMBERS_PENDING from the database

			SpecialSql8(Access) = "DROP TABLE " & strMemberTablePrefix & "MEMBERS_PENDING"
			SpecialSql8(SQL6) = "DROP TABLE " & strMemberTablePrefix & "MEMBERS_PENDING"
			SpecialSql8(SQL7) = "DROP TABLE " & strMemberTablePrefix & "MEMBERS_PENDING"
			SpecialSql8(MySql) = "DROP TABLE IF EXISTS " & strMemberTablePrefix & "MEMBERS_PENDING"

	 		strOkMessage = "Table MEMBERS_PENDING has been dropped"

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Add Table MEMBERS_PENDING to the database

			SpecialSql8(Access) = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ( "
			SpecialSql8(Access) = SpecialSql8(Access) & "MEMBER_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_STATUS smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_NAME text (75) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_USERNAME text (150) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_PASSWORD text (65) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_EMAIL text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_COUNTRY text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_HOMEPAGE text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_SIG memo NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_VIEW_SIG smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_SIG_DEFAULT smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_DEFAULT_VIEW int NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LEVEL smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_AIM text (150) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_ICQ text (150) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_MSN text (150) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_YAHOO text (150) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_POSTS int NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_DATE text (14) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LASTHEREDATE text (14) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LASTPOSTDATE text (14) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_TITLE text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_SUBSCRIPTION smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_HIDE_EMAIL smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_RECEIVE_EMAIL smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LAST_IP text (15) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_IP text (15) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_FIRSTNAME text (100) NULL ,"
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LASTNAME text (100) NULL ,"
			SpecialSql8(Access) = SpecialSql8(Access) & "M_OCCUPATION text (255) NULL ,"
			SpecialSql8(Access) = SpecialSql8(Access) & "M_SEX text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_AGE text (3) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_DOB text (8) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_HOBBIES memo NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LNEWS memo NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_QUOTE memo NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_BIO memo NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_MARSTATUS text (100) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LINK1 text (255) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_LINK2 text (255) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_CITY text (100) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_STATE text (100) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_PHOTO_URL text (255) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_KEY text (32) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_NEWEMAIL text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_PWKEY text (32) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_APPROVE smallint NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "M_SHA256 smallint NULL ) "

			SpecialSql8(SQL6) = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ( "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "MEMBER_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_STATUS smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_NAME varchar (75) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_USERNAME varchar (150) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_PASSWORD varchar (65) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_EMAIL varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_COUNTRY varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_HOMEPAGE varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_SIG text NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_VIEW_SIG smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_SIG_DEFAULT smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_DEFAULT_VIEW int NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LEVEL smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_AIM varchar (150) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_ICQ varchar (150) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_MSN varchar (150) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_YAHOO varchar (150) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_POSTS int NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_DATE varchar (14) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LASTHEREDATE varchar (14) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LASTPOSTDATE varchar (14) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_TITLE varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_SUBSCRIPTION smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_HIDE_EMAIL smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_RECEIVE_EMAIL smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LAST_IP varchar (15) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_IP varchar (15) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_FIRSTNAME varchar (100) NULL ,"
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LASTNAME varchar (100) NULL ,"
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_OCCUPATION varchar (255) NULL ,"
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_SEX varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_AGE varchar (3) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_DOB varchar (8) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_HOBBIES text NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LNEWS text NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_QUOTE text NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_BIO text NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_MARSTATUS varchar (100) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LINK1 varchar (255) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_LINK2 varchar (255) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_CITY varchar (100) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_STATE varchar (100) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_PHOTO_URL varchar (255) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_KEY varchar (32) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_NEWEMAIL varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_PWKEY varchar (32) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_APPROVE smallint NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "M_SHA256 smallint NULL ) "

			SpecialSql8(SQL7) = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ( "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "MEMBER_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_STATUS smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_NAME nvarchar (75) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_USERNAME nvarchar (150) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_PASSWORD nvarchar (65) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_EMAIL nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_COUNTRY nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_HOMEPAGE nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_SIG ntext NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_VIEW_SIG smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_SIG_DEFAULT smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_DEFAULT_VIEW int NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LEVEL smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_AIM nvarchar (150) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_ICQ nvarchar (150) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_MSN nvarchar (150) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_YAHOO nvarchar (150) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_POSTS int NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_DATE nvarchar (14) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LASTHEREDATE nvarchar (14) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LASTPOSTDATE nvarchar (14) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_TITLE nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_SUBSCRIPTION smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_HIDE_EMAIL smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_RECEIVE_EMAIL smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LAST_IP nvarchar (15) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_IP nvarchar (15) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_FIRSTNAME nvarchar (100) NULL ,"
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LASTNAME nvarchar (100) NULL ,"
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_OCCUPATION nvarchar (255) NULL ,"
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_SEX nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_AGE nvarchar (3) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_DOB nvarchar (8) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_HOBBIES ntext NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LNEWS ntext NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_QUOTE ntext NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_BIO ntext NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_MARSTATUS nvarchar (100) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LINK1 nvarchar (255) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_LINK2 nvarchar (255) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_CITY nvarchar (100) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_STATE nvarchar (100) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_PHOTO_URL nvarchar (255) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_KEY nvarchar (32) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_NEWEMAIL nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_PWKEY nvarchar (32) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_APPROVE smallint NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "M_SHA256 smallint NULL ) "

			SpecialSql8(MySql) = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ( "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "MEMBER_ID int (11) NOT NULL auto_increment , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_STATUS smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_NAME varchar (75) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_USERNAME varchar (150) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_PASSWORD varchar (65) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_EMAIL varchar (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_COUNTRY varchar (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_HOMEPAGE varchar (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_SIG text NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_VIEW_SIG smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_SIG_DEFAULT smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_DEFAULT_VIEW int (11) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LEVEL smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_AIM varchar (150) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_ICQ varchar (150) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_MSN varchar (150) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_YAHOO varchar (150) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_POSTS int (11) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_DATE varchar (14) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LASTHEREDATE varchar (14) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LASTPOSTDATE varchar (14) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_TITLE varchar (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_SUBSCRIPTION smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_HIDE_EMAIL smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_RECEIVE_EMAIL smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LAST_IP varchar (15) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_IP varchar (15) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_FIRSTNAME varchar (100) NULL ,"
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LASTNAME varchar (100) NULL ,"
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_OCCUPATION varchar (255) NULL ,"
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_SEX varchar (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_AGE varchar (3) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_DOB varchar (8) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_HOBBIES text NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LNEWS text NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_QUOTE text NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_BIO text NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_MARSTATUS varchar (100) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LINK1 varchar (255) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_LINK2 varchar (255) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_CITY varchar (100) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_STATE varchar (100) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_PHOTO_URL varchar (255) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_KEY varchar (32) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_NEWEMAIL varchar (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_PWKEY varchar (32) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_APPROVE smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "M_SHA256 smallint (6) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "PRIMARY KEY (MEMBER_ID) ) "

	 		strOkMessage = "Table MEMBERS_PENDING created "

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Add Table GROUP_NAMES to the database

			SpecialSql8(Access) = "CREATE TABLE " & strTablePrefix & "GROUP_NAMES ( "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_NAME text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_DESCRIPTION text (255) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_ICON text (255) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_IMAGE text (255) NULL )"

			SpecialSql8(SQL6) = "CREATE TABLE " & strTablePrefix & "GROUP_NAMES ( "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_NAME varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_DESCRIPTION varchar (255) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_ICON varchar (255) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_IMAGE varchar (255) NULL )"

			SpecialSql8(SQL7) = "CREATE TABLE " & strTablePrefix & "GROUP_NAMES ( "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_NAME nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_DESCRIPTION nvarchar (255) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_ICON nvarchar (255) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_IMAGE nvarchar (255) NULL )"

			SpecialSql8(MySql) = "CREATE TABLE " & strTablePrefix & "GROUP_NAMES ( "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_ID int (11) NOT NULL auto_increment , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_NAME VARCHAR (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_DESCRIPTION VARCHAR (255) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_ICON VARCHAR (255) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_IMAGE VARCHAR (255) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "PRIMARY KEY (GROUP_ID)) "

	 		strOkMessage = "Table GROUP_NAMES created "

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Add Table GROUP to the database

			SpecialSql8(Access) = "CREATE TABLE " & strTablePrefix & "GROUPS ( "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_KEY COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_ID int NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "GROUP_CATID int NULL )"

			SpecialSql8(SQL6) = "CREATE TABLE " & strTablePrefix & "GROUPS ( "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_KEY int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_ID int NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "GROUP_CATID int NULL )"

			SpecialSql8(SQL7) = "CREATE TABLE " & strTablePrefix & "GROUPS ( "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_KEY int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_ID int NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "GROUP_CATID int NULL )"
                                                                                      
			SpecialSql8(MySql) = "CREATE TABLE " & strTablePrefix & "GROUPS ( "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_KEY int (11) NOT NULL auto_increment , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_ID int (11) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "GROUP_CATID int (11) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "PRIMARY KEY (GROUP_KEY)) "

	 		strOkMessage = "Table GROUPS created "

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Check that there are no records in the GROUP_NAMES table

			strSql7 = "SELECT GROUP_ID FROM " & strTablePrefix & "GROUP_NAMES "

			set rs7 = my_Conn.Execute(strSql7)

			if rs7.EOF then
				'## Add Default Group 1 to the GROUP_NAMES table

				SpecialSql8(Access) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(Access) = SpecialSql8(Access) & "VALUES ('All Categories you have access to','All Categories you have access to')"

				SpecialSql8(SQL6) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(SQL6) = SpecialSql8(SQL6) & "VALUES ('All Categories you have access to','All Categories you have access to')"

				SpecialSql8(SQL7) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(SQL7) = SpecialSql8(SQL7) & "VALUES ('All Categories you have access to','All Categories you have access to')"
                                                                                      
				SpecialSql8(MySql) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(MySql) = SpecialSql8(MySql) & "VALUES ('All Categories you have access to','All Categories you have access to')"

		 		strOkMessage = "New Record inserted into GROUP_NAMES table"

				call SpecialUpdates(SpecialSql8, strOkMessage)

				Response.Flush

				'## Add Default Group 2 to the GROUP_NAMES table

				SpecialSql8(Access) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(Access) = SpecialSql8(Access) & "VALUES ('Default Categories','Default Categories')"

				SpecialSql8(SQL6) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(SQL6) = SpecialSql8(SQL6) & "VALUES ('Default Categories','Default Categories')"

				SpecialSql8(SQL7) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(SQL7) = SpecialSql8(SQL7) & "VALUES ('Default Categories','Default Categories')"
                                                                                      
				SpecialSql8(MySql) = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) "
				SpecialSql8(MySql) = SpecialSql8(MySql) & "VALUES ('Default Categories','Default Categories')"

	 			strOkMessage = "New Record inserted into GROUP_NAMES table"

				call SpecialUpdates(SpecialSql8, strOkMessage)

				Response.Flush
			end if
			rs7.close
			set rs7 = nothing

			'## Add Table NAMEFILTER to the database

			SpecialSql8(Access) = "CREATE TABLE " & strFilterTablePrefix & "NAMEFILTER ( "
			SpecialSql8(Access) = SpecialSql8(Access) & "N_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql8(Access) = SpecialSql8(Access) & "N_NAME text (75) NULL )"

			SpecialSql8(SQL6) = "CREATE TABLE " & strFilterTablePrefix & "NAMEFILTER ( "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "N_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "N_NAME varchar (75) NULL )"

			SpecialSql8(SQL7) = "CREATE TABLE " & strFilterTablePrefix & "NAMEFILTER ( "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "N_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "N_NAME nvarchar (75) NULL )"

			SpecialSql8(MySql) = "CREATE TABLE " & strFilterTablePrefix & "NAMEFILTER ( "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "N_ID int (11) NOT NULL auto_increment , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "N_NAME VARCHAR (75) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "PRIMARY KEY (N_ID)) "			

	 		strOkMessage = "Table NAMEFILTER created "

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Add Table BADWORDS to the database

			SpecialSql8(Access) = "CREATE TABLE " & strFilterTablePrefix & "BADWORDS ( "
			SpecialSql8(Access) = SpecialSql8(Access) & "B_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql8(Access) = SpecialSql8(Access) & "B_BADWORD text (50) NULL , "
			SpecialSql8(Access) = SpecialSql8(Access) & "B_REPLACE text (50) NULL )"

			SpecialSql8(SQL6) = "CREATE TABLE " & strFilterTablePrefix & "BADWORDS ( "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "B_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "B_BADWORD varchar (50) NULL , "
			SpecialSql8(SQL6) = SpecialSql8(SQL6) & "B_REPLACE varchar (50) NULL )"

			SpecialSql8(SQL7) = "CREATE TABLE " & strFilterTablePrefix & "BADWORDS ( "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "B_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "B_BADWORD nvarchar (50) NULL , "
			SpecialSql8(SQL7) = SpecialSql8(SQL7) & "B_REPLACE nvarchar (50) NULL )"

			SpecialSql8(MySql) = "CREATE TABLE " & strFilterTablePrefix & "BADWORDS ( "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "B_ID int (11) NOT NULL auto_increment , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "B_BADWORD VARCHAR (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "B_REPLACE VARCHAR (50) NULL , "
			SpecialSql8(MySql) = SpecialSql8(MySql) & "PRIMARY KEY (B_ID)) "			

	 		strOkMessage = "Table BADWORDS created "

			call SpecialUpdates(SpecialSql8, strOkMessage)

			Response.Flush

			'## Add current Badwords to the BADWORDS table

			if strBadWords = "" then
				strSql = "SELECT C_VALUE "
				strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
				strSql = strSql & " WHERE C_VARIABLE = 'strBadWords' "

				set rsConfig = my_Conn.Execute (StrSql)

				if not rsConfig.EOF then
					strBadWords = rsConfig("C_VALUE")
				end if

				set rsConfig = nothing
			end if

			bwords = split(strBadWords, "|")
			for b = 0 to ubound(bwords)
				if bwords(b) <> "" then
					txtBadWord = bwords(b)
					txtReplace = Replace(txtBadWord, bwords(b), string(len(bwords(b)),"*"), 1,-1,1)
					strDummy = SetBadWordValue(0,txtBadWord,txtReplace)
				end if
			next

			if strDummy = "added" then
				Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Existing Badwords Added to Badwords Table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("</table>" & vbNewLine)
			end if

			'## Add the new config-values to the database

			strDummy = SetConfigValue(0,"STRARCHIVESTATE","1")
			strDummy = SetConfigValue(0,"STRFLOODCHECK","1")
			strDummy = SetConfigValue(0,"STRFLOODCHECKTIME","-60")
			strDummy = SetConfigValue(0,"STREMAILVAL","0")
			strDummy = SetConfigValue(0,"STRPAGEBGIMAGEURL"," ")
			strDummy = SetConfigValue(0,"STRIMAGEURL"," ")
			strDummy = SetConfigValue(0,"STRJUMPLASTPOST","0")
			strDummy = SetConfigValue(0,"STRSTICKYTOPIC","0")
			strDummy = SetConfigValue(0,"STRDSIGNATURES","0")
			strDummy = SetConfigValue(0,"STRSHOWSENDTOFRIEND","1")
			strDummy = SetConfigValue(0,"STRSHOWPRINTERFRIENDLY","1")
			strDummy = SetConfigValue(0,"STRPROHIBITNEWMEMBERS","0")
			strDummy = SetConfigValue(0,"STRREQUIREREG","0")
			strDummy = SetConfigValue(0,"STRRESTRICTREG","0")
			strDummy = SetConfigValue(0,"STRHILITEFONTCOLOR","red")
			strDummy = SetConfigValue(0,"STRSEARCHHILITECOLOR","yellow")
			strDummy = SetConfigValue(0,"STRGROUPCATEGORIES","0")
			strDummy = SetConfigValue(0,"STRACTIVETEXTDECORATION","underline")
			strDummy = SetConfigValue(0,"STRFORUMLINKTEXTDECORATION","underline")
			strDummy = SetConfigValue(0,"STRFORUMVISITEDLINKCOLOR","blue")
			strDummy = SetConfigValue(0,"STRFORUMVISITEDTEXTDECORATION","underline")
			strDummy = SetConfigValue(0,"STRFORUMACTIVELINKCOLOR","red")
			strDummy = SetConfigValue(0,"STRFORUMACTIVETEXTDECORATION","underline")
			strDummy = SetConfigValue(0,"STRFORUMHOVERFONTCOLOR","red")
			strDummy = SetConfigValue(0,"STRFORUMHOVERTEXTDECORATION","underline")
			strDummy = SetConfigValue(0,"STRSHOWTIMER","0")
			strDummy = SetConfigValue(0,"STRTIMERPHRASE","This page was generated in [TIMER] seconds.")
			strDummy = SetConfigValue(0,"STRSHOWFORMATBUTTONS","1")
			strDummy = SetConfigValue(0,"STRSHOWSMILIESTABLE","1")
			strDummy = SetConfigValue(0,"STRSHOWQUICKREPLY","0")
			strDummy = SetConfigValue(0,"STRAGEDOB","0")

			Response.Flush

			set Server2 = Server
			Server2.ScriptTimeout = 10000

			'## Encrypt existing Passwords
			strSql = "SELECT MEMBER_ID, M_PASSWORD "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE (M_SHA256 IS NULL) "
			strSql = strSql & " ORDER BY MEMBER_ID "

			set rsenc =  Server.CreateObject("ADODB.RecordSet")
			rsenc.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rsenc.EOF then
				recMemberCount = ""
			else
				allMemberData = rsenc.GetRows(adGetRowsRest)
				recMemberCount = UBound(allMemberData,2)
			end if

			rsenc.close
			set rsenc = nothing

			if recMemberCount <> "" then
				mMEMBER_ID = 0
				mM_PASSWORD = 1

				ipwd = 0
				ipwdt = 0
				Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Processing: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Encrypting Existing Passwords (Please Wait...)</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("</table>" & vbNewLine)
				Response.Flush
				for iMember = 0 to recMemberCount
					mMemberID = allMemberData(mMEMBER_ID, iMember)
					mMemberPassword = allMemberData(mM_PASSWORD, iMember)

					ipwd = ipwd + 1
					ipwdt = ipwdt + 1
					strEncodedPassword = sha256("" & mMemberPassword)

					strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " SET M_PASSWORD = '" & strEncodedPassword & "'"
					strSql = strSql & " ,   M_SHA256 = 1 "
					strSql = strSql & " WHERE MEMBER_ID = " & mMemberID

					on error resume next
					my_Conn.Errors.Clear
					Err.Clear
					my_Conn.Execute(strSql)
					Response.Write "."
					if ipwd = 100 then 
						Response.Write(" <font face=""Verdana, Arial, Helvetica"" size=""1"">(" & ipwdt & " records processed)</font><br />" & vbNewLine)
						Response.Flush
						ipwd = 0
					end if
				next
				Response.Write("<br /> <font face=""Verdana, Arial, Helvetica"" size=""1"">(" & ipwdt & " total records processed)</font>" & vbNewLine)
				Response.Flush
				Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Existing Passwords Encrypted</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("</table>" & vbNewLine)
			end if

			Response.Flush

			'## update the version info...

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if

'###########################################################################
'##
'## Setup for update 9 / to version 3.4.01
'##
'###########################################################################

		if (OldVersion <= 9) then

			'## Add the new config-values to the database

			strDummy = SetConfigValue(0,"STRUSERNAMEFILTER","0")

			Response.Flush

			'## update the version info...

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if

'###########################################################################
'##
'## Setup for update 10 / to version 3.4.02, 3.4.03 and 3.4.04
'##
'###########################################################################

		if (OldVersion <= 10) then

			Dim SpecialSql10(4)

			'## Update T_STICKY for existing Topics in TOPICS table

			SpecialSql10(Access) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql10(SQL6) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql10(SQL7) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"
			SpecialSql10(MySql) = "UPDATE " & strTablePrefix & "TOPICS SET T_STICKY = 0 WHERE (T_STICKY IS NULL)"

			strOkMessage = "T_STICKY field value updated in the TOPICS table"

			call SpecialUpdates(SpecialSql10, strOkMessage)
			
			'## Change SQL 7 T_MESSAGE, T_SUBJECT and R_MESSAGE fields to unicode versions correcting an old bug
			if (strDBType = "sqlserver") and (strSQL_Server = "SQL7") then
				'Dim SpecialSQL12(4)
				'## Change T_MESSAGE to a NTEXT in A_TOPICS table and TOPICS table
				strOkMessage = "T_MESSAGE Field has been changed"

		 		SpecialSQL10(SQL7) = "ALTER TABLE " & strTablePrefix & "A_TOPICS ALTER COLUMN T_MESSAGE NTEXT NULL "
				call SpecialUpdates(SpecialSQL10, strOkMessage)
				Response.Flush

		 		SpecialSQL10(SQL7) = "ALTER TABLE " & strTablePrefix & "TOPICS ALTER COLUMN T_MESSAGE NTEXT NULL "
				call SpecialUpdates(SpecialSQL10, strOkMessage)
				Response.Flush

				'## Change T_SUBJECT to a NVARCHAR in A_TOPICS table and TOPICS table
				strOkMessage = "T_SUBJECT Field has been changed"

		 		SpecialSQL10(SQL7) = "ALTER TABLE " & strTablePrefix & "A_TOPICS ALTER COLUMN T_SUBJECT NVARCHAR (100) NULL "
				call SpecialUpdates(SpecialSQL10, strOkMessage)
				Response.Flush

		 		SpecialSQL10(SQL7) = "ALTER TABLE " & strTablePrefix & "TOPICS ALTER COLUMN T_SUBJECT NVARCHAR (100) NULL "
				call SpecialUpdates(SpecialSQL10, strOkMessage)
				Response.Flush

				'## Change R_MESSAGE to a NTEXT in A_REPLY table and REPLY table
				strOkMessage = "R_MESSAGE Field has been changed"

		 		SpecialSQL10(SQL7) = "ALTER TABLE " & strTablePrefix & "A_REPLY ALTER COLUMN R_MESSAGE NTEXT NULL "
				call SpecialUpdates(SpecialSQL10, strOkMessage)
				Response.Flush

		 		SpecialSQL10(SQL7) = "ALTER TABLE " & strTablePrefix & "REPLY ALTER COLUMN R_MESSAGE NTEXT NULL "
				call SpecialUpdates(SpecialSQL10, strOkMessage)
				Response.Flush
			end if

			Response.Flush

			'## update the version info...

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if

'###########################################################################
'##
'## Setup for update 11 / to version 3.4.06
'##
'###########################################################################

		if (OldVersion <= 11) then
			
			'## Add the new config-values to the database

			strDummy = SetConfigValue(0,"STRMINAGE","13")

			Response.Flush

			'## update the version info...

			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if


'###########################################################################
'##
'## Setup for update 12 / to version 3.4.07
'##
'###########################################################################

		if (OldVersion <= 12) then
			
			Dim SpecialSql12(4)

			'## Change F_IP, M_LAST_IP, M_IP, R_IP, T_IP fields to VARCHAR 50 to hold IPv6 ip addresses.
			strOkMessage = "F_IP field in " & strTablePrefix & "FORUM has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "FORUM ALTER COLUMN F_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "FORUM ALTER COLUMN F_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "FORUM ALTER COLUMN F_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "M_LAST_IP field in " & strTablePrefix & "MEMBERS table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "MEMBERS ALTER COLUMN M_LAST_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "MEMBERS ALTER COLUMN M_LAST_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "MEMBERS ALTER COLUMN M_LAST_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "MEMBERS MODIFY M_LAST_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "M_IP field in " & strTablePrefix & "MEMBERS table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "MEMBERS ALTER COLUMN M_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "MEMBERS ALTER COLUMN M_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "MEMBERS ALTER COLUMN M_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "MEMBERS MODIFY M_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "M_LAST_IP field in " & strTablePrefix & "MEMBERS_PENDING table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ALTER COLUMN M_LAST_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ALTER COLUMN M_LAST_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ALTER COLUMN M_LAST_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING MODIFY M_LAST_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "M_IP field in " & strTablePrefix & "MEMBERS_PENDING table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ALTER COLUMN M_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ALTER COLUMN M_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ALTER COLUMN M_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING MODIFY M_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "R_IP field in " & strTablePrefix & "REPLY table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "REPLY ALTER COLUMN R_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "REPLY ALTER COLUMN R_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "REPLY ALTER COLUMN R_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "REPLY MODIFY R_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "R_IP field in " & strTablePrefix & "A_REPLY table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "A_REPLY ALTER COLUMN R_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "A_REPLY ALTER COLUMN R_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "A_REPLY ALTER COLUMN R_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "A_REPLY MODIFY R_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "T_IP field in " & strTablePrefix & "TOPICS table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "TOPICS ALTER COLUMN T_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "TOPICS ALTER COLUMN T_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "TOPICS ALTER COLUMN T_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "TOPICS MODIFY T_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "T_IP field in " & strTablePrefix & "A_TOPICS table has been changed"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strTablePrefix & "A_TOPICS ALTER COLUMN T_IP TEXT (50) "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strTablePrefix & "A_TOPICS ALTER COLUMN T_IP VARCHAR (50) '000.000.000.000' "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strTablePrefix & "A_TOPICS ALTER COLUMN T_IP NVARCHAR (50) '000.000.000.000' "
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "A_TOPICS MODIFY T_IP VARCHAR (50) DEFAULT '000.000.000.000' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			
			'## Change T_REPLIES, F_TOPICS, F_COUNT, F_A_TOPICS, F_A_COUNT to DOUBLE type for MySQL
			strOkMessage = "T_REPLIES field in " & strTablePrefix & "TOPICS table has been changed"
			
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "TOPICS MODIFY T_REPLIES DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "T_REPLIES field in " & strTablePrefix & "A_TOPICS table has been changed"
			
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "A_TOPICS MODIFY T_REPLIES DOUBLE "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "F_TOPICS field in " & strTablePrefix & "FORUM table has been changed"
			
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_TOPICS DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "F_A_TOPICS field in " & strTablePrefix & "FORUM table has been changed"
			
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_A_TOPICS DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "F_COUNT field in " & strTablePrefix & "FORUM table has been changed"
			
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_COUNT DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			strOkMessage = "F_A_COUNT field in " & strTablePrefix & "FORUM table has been changed"
			
			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_A_COUNT DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			
			'## Add the M_ALLOWEMAIL field to MEMBERS and MEMBERS_PENDING
			strOkMessage = "M_ALLOWEMAIL has been added to " & strMemberTablePrefix & "MEMBERS"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0"
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			'Update defaults?
			
			strOkMessage = "M_ALLOWEMAIL has been added to " & strMemberTablePrefix & "MEMBERS_PENDING"
			
			SpecialSQL12(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL6) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL7) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(MySql) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0"
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			'Update defaults?
			
			'## Add Table SPAM_MAIL to the database
	 		strOkMessage = "Table " & strFilterTablePrefix & "SPAM_MAIL created"

			SpecialSQL12(Access) = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(Access) = SpecialSQL12(Access) & "SPAM_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSQL12(Access) = SpecialSQL12(Access) & "SPAM_SERVER text (255) NULL )"

			SpecialSQL12(SQL6) = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(SQL6) = SpecialSQL12(SQL6) & "SPAM_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSQL12(SQL6) = SpecialSQL12(SQL6) & "SPAM_SERVER varchar (255) NULL )"

			SpecialSQL12(SQL7) = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(SQL7) = SpecialSQL12(SQL7) & "SPAM_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSQL12(SQL7) = SpecialSQL12(SQL7) & "SPAM_SERVER nvarchar (255) NULL )"

			SpecialSQL12(MySql) = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(MySql) = SpecialSQL12(MySql) & "SPAM_ID int (11) NOT NULL auto_increment , "
			SpecialSQL12(MySql) = SpecialSQL12(MySql) & "SPAM_SERVER VARCHAR (255) NULL , "
			SpecialSQL12(MySql) = SpecialSQL12(MySql) & "PRIMARY KEY (SPAM_ID)) "			
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			
			'## Add the new config-values to the database
			strDummy = SetConfigValue(0,"STRREQAIM","0")
			strDummy = SetConfigValue(0,"STRREQICQ","0")
			strDummy = SetConfigValue(0,"STRREQMSN","0")
			strDummy = SetConfigValue(0,"STRREQYAHOO","0")
			strDummy = SetConfigValue(0,"STRREQFULLNAME","0")
			strDummy = SetConfigValue(0,"STRREQPICTURE","0")
			strDummy = SetConfigValue(0,"STRREQCITY","0")
			strDummy = SetConfigValue(0,"STRREQSTATE","0")
			strDummy = SetConfigValue(0,"STRREQAGE","0")
			strDummy = SetConfigValue(0,"STRREQAGEDOB","0")
			strDummy = SetConfigValue(0,"STRREQHOMEPAGE","0")
			strDummy = SetConfigValue(0,"STRREQCOUNTRY","0")
			strDummy = SetConfigValue(0,"STRREQOCCUPATION","0")
			strDummy = SetConfigValue(0,"STRREQBIO","0")
			strDummy = SetConfigValue(0,"STRREQHOBBIES","0")
			strDummy = SetConfigValue(0,"STRREQLNEWS","0")
			strDummy = SetConfigValue(0,"STRREQQUOTE","0")
			strDummy = SetConfigValue(0,"STRREQMARSTATUS","0")
			strDummy = SetConfigValue(0,"STRREQFAVLINKS","0")
			strDummy = SetConfigValue(0,"INTMAXPOSTSTOEMAIL","10")
			strDummy = SetConfigValue(0,"STRNOMAXPOSTSTOEMAIL","You do not have enough posts to email other members. If you feel that you have received this message in error, please contact the forum administrator.")
			strDummy = SetConfigValue(0,"STRFILTEREMAILADDRESSES","0")
						
			'## update the version info...
			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if


'###########################################################################
'##
'## Setup for update 13 / to version 3.5.00 - suggested
'##
'###########################################################################

		if (OldVersion <= 13 and 1=2) then
			
			Dim SpecialSql13(4)
			
			'## Add the new table for the custom policy MOD
	 		strOkMessage = "Table " & strFilterTablePrefix & "CUSTOM_POLICY created"
			
			SpecialSql13(Access) = "CREATE TABLE " & strTablePrefix & "CUSTOM_POLICY ( "
			SpecialSql13(Access) = SpecialSql13(Access) & "CP_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql13(Access) = SpecialSql13(Access) & "CP_MODE smallint NOT NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "CP_CONTENT memo NULL )"
			
			SpecialSql13(SQL6) = "CREATE TABLE " & strTablePrefix & "CUSTOM_POLICY ( "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "CP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "CP_MODE smallint NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "CP_CONTENT text NULL )"
			
			SpecialSql13(SQL7) = "CREATE TABLE " & strTablePrefix & "CUSTOM_POLICY ( "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "CP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "CP_MODE smallint NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "CP_CONTENT ntext NULL )"
			
			SpecialSql13(MySql) = "CREATE TABLE " & strTablePrefix & "CUSTOM_POLICY ( "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "CP_ID int (11) NOT NULL auto_increment , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "CP_MODE smallint (6) NOT NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "CP_CONTENT text NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "PRIMARY KEY (CP_ID)) "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			'## Add the defaults for the custom policy MOD
			strSql = "INSERT INTO " & strTablePrefix & "CUSTOM_POLICY "
			strSql = strSql & " (CP_MODE, CP_CONTENT)"
			strSql = strSql & " VALUES (0,'Your custom policy here')"
			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)
			on error goto 0
			
			
			'## Add the new tables for the custom FAQ MOD
	 		strOkMessage = "Table " & strFilterTablePrefix & "FAQ created"
			
			SpecialSql13(Access) = "CREATE TABLE " & strTablePrefix & "FAQ ( "
			SpecialSql13(Access) = SpecialSql13(Access) & "F_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql13(Access) = SpecialSql13(Access) & "F_FAQ_QUESTION text (255) NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "F_FAQ_ANSWER memo NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "F_FAQ_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(Access) = SpecialSql13(Access) & "F_FAQ_CATEGORY int NULL DEFAULT 1 , "
			SpecialSql13(Access) = SpecialSql13(Access) & "F_FAQ_TYPE int NULL DEFAULT 0)"
			
			SpecialSql13(SQL6) = "CREATE TABLE " & strTablePrefix & "FAQ ( "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "F_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "F_FAQ_QUESTION varchar (255) NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "F_FAQ_ANSWER text NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "F_FAQ_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "F_FAQ_CATEGORY int NULL DEFAULT 1 , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "F_FAQ_TYPE int NULL DEFAULT 0)"
			
			SpecialSql13(SQL7) = "CREATE TABLE " & strTablePrefix & "FAQ ( "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "F_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "F_FAQ_QUESTION nvarchar (255) NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "F_FAQ_ANSWER ntext NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "F_FAQ_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "F_FAQ_CATEGORY int NULL DEFAULT 1 , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "F_FAQ_TYPE int NULL DEFAULT 0)"
			
			SpecialSql13(MySql) = "CREATE TABLE " & strTablePrefix & "FAQ ( "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "F_ID int (11) NOT NULL auto_increment , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "F_FAQ_QUESTION varchar (255) NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "F_FAQ_ANSWER text NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "F_FAQ_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "F_FAQ_CATEGORY int NULL DEFAULT 1 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "F_FAQ_TYPE int NULL DEFAULT 0 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "PRIMARY KEY (F_ID)) "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			strOkMessage = "Table " & strFilterTablePrefix & "FAQ_CATEGORY created"
			
			SpecialSql13(Access) = "CREATE TABLE " & strTablePrefix & "FAQ_CATEGORY ( "
			SpecialSql13(Access) = SpecialSql13(Access) & "FCAT_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql13(Access) = SpecialSql13(Access) & "FCAT_TITLE text (100) NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "FCAT_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(Access) = SpecialSql13(Access) & "FCAT_LEVEL int NULL)"
			
			SpecialSql13(SQL6) = "CREATE TABLE " & strTablePrefix & "FAQ_CATEGORY ( "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "FCAT_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "FCAT_TITLE varchar (100) NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "FCAT_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "FCAT_LEVEL int NULL)"
			
			SpecialSql13(SQL7) = "CREATE TABLE " & strTablePrefix & "FAQ_CATEGORY ( "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "FCAT_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "FCAT_TITLE nvarchar (100) NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "FCAT_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "FCAT_LEVEL int NULL)"
			
			SpecialSql13(MySql) = "CREATE TABLE " & strTablePrefix & "FAQ_CATEGORY ( "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "FCAT_ID int (11) NOT NULL auto_increment , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "FCAT_TITLE varchar (100) NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "FCAT_ORDER int NULL DEFAULT 1 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "FCAT_LEVEL int NULL DEFAULT 1 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "PRIMARY KEY (FCAT_ID)) "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			'## Add the defaults for the custom FAQ MOD
			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			strSql = "INSERT INTO " & strTablePrefix & "FAQ_CATEGORY "
			strSql = strSql & "(FCAT_TITLE,FCAT_ORDER,FCAT_LEVEL) VALUES ('General',1,0)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Do I have to register?','',1,1,1)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('How can I use smilies and images?','',1,1,2)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I add a hyperlink to my messages?','',1,1,3)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I change the format of my text?','',1,1,4)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What are Moderators?','',1,1,5)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Are cookies used?','',1,1,6)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What are active topics?','',1,1,7)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I edit my own posts?','',1,1,8)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I attach files?','',1,1,9)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I search?','',1,1,10)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I edit my profile?','',1,1,11)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I attach my own signature to my posts?','',1,1,12)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Are there any censor features?','',1,1,13)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What do I do if I forget my Password?','',1,1,14)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I be notified by e-mail when there are new posts?','',1,1,15)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What does it mean if a forum has Moderation enabled?','',1,1,16)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What is COPPA?','',1,1,17)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Where can I get my own copy of this Forum?','',1,1,18)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What is RSS?','',1,1,19)"
			my_Conn.Execute (strSql)
			strSql = "INSERT INTO " & strTablePrefix & "FAQ "
			strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('How do I access an RSS Feed?','',1,1,20)"
			my_Conn.Execute (strSql)
			on error goto 0
			
			
			'## Add the new tables for the UserGroup MOD
			strOkMessage = "Table " & strFilterTablePrefix & "USERGROUPS created"
			
			SpecialSql13(Access) = "CREATE TABLE " & strTablePrefix & "USERGROUPS ( "
			SpecialSql13(Access) = SpecialSql13(Access) & "USERGROUP_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSql13(Access) = SpecialSql13(Access) & "USERGROUP_NAME varchar(100) NOT NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "USERGROUP_DESC memo NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "USERGROUP_TYPE int NOT NULL DEFAULT 0 , "
			SpecialSql13(Access) = SpecialSql13(Access) & "AUTOJOIN int NOT NULL DEFAULT 0 , "
			SpecialSql13(Access) = SpecialSql13(Access) & "MEM_HIDE int NOT NULL DEFAULT 0 , "
			SpecialSql13(Access) = SpecialSql13(Access) & "MOD_HIDE int NOT NULL DEFAULT 0)"
			
			SpecialSql13(SQL6) = "CREATE TABLE " & strTablePrefix & "USERGROUPS ( "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "USERGROUP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "USERGROUP_NAME varchar(100) NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "USERGROUP_DESC text NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "USERGROUP_TYPE int NOT NULL DEFAULT 0 , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "AUTOJOIN int NOT NULL DEFAULT 0 , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "MEM_HIDE int NOT NULL DEFAULT 0 , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "MOD_HIDE int NOT NULL DEFAULT 0)"
			
			SpecialSql13(SQL7) = "CREATE TABLE " & strTablePrefix & "USERGROUPS ( "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "USERGROUP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "USERGROUP_NAME nvarchar(100) NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "USERGROUP_DESC ntext NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "USERGROUP_TYPE int NOT NULL DEFAULT 0 , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "AUTOJOIN int NOT NULL DEFAULT 0 , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "MEM_HIDE int NOT NULL DEFAULT 0 , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "MOD_HIDE int NOT NULL DEFAULT 0)"
			
			SpecialSql13(MySql) = "CREATE TABLE " & strTablePrefix & "USERGROUPS ( "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "USERGROUP_ID int (11) NOT NULL auto_increment , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "USERGROUP_NAME varchar(100) NOT NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "USERGROUP_DESC text NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "USERGROUP_TYPE int NOT NULL DEFAULT 0 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "AUTOJOIN int NOT NULL DEFAULT 0 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "MEM_HIDE int NOT NULL DEFAULT 0 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "MOD_HIDE int NOT NULL DEFAULT 0 , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "PRIMARY KEY (USERGROUP_ID)) "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			strOkMessage = "Table " & strFilterTablePrefix & "USERGROUP_USERS created"
			
			SpecialSql13(Access) = "CREATE TABLE " & strTablePrefix & "USERGROUP_USERS ( "
			SpecialSql13(Access) = SpecialSql13(Access) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "MEMBER_ID int NOT NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "MEMBER_TYPE int NOT NULL DEFAULT 1)"
			
			SpecialSql13(SQL6) = "CREATE TABLE " & strTablePrefix & "USERGROUP_USERS ( "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "MEMBER_ID int NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "MEMBER_TYPE int NOT NULL DEFAULT 1)"
			
			SpecialSql13(SQL7) = "CREATE TABLE " & strTablePrefix & "USERGROUP_USERS ( "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "MEMBER_ID int NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "MEMBER_TYPE int NOT NULL DEFAULT 1)"
			
			SpecialSql13(MySql) = "CREATE TABLE " & strTablePrefix & "USERGROUP_USERS ( "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "MEMBER_ID int NOT NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "MEMBER_TYPE int NOT NULL DEFAULT 1)"
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			strOkMessage = "Table " & strFilterTablePrefix & "ALLOWED_USERGROUPS created"
			
			SpecialSql13(Access) = "CREATE TABLE " & strTablePrefix & "ALLOWED_USERGROUPS ( "
			SpecialSql13(Access) = SpecialSql13(Access) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "FORUM_ID int NOT NULL , "
			SpecialSql13(Access) = SpecialSql13(Access) & "PERMS int NOT NULL DEFAULT 0)"
			
			SpecialSql13(SQL6) = "CREATE TABLE " & strTablePrefix & "ALLOWED_USERGROUPS ( "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "FORUM_ID int NOT NULL , "
			SpecialSql13(SQL6) = SpecialSql13(SQL6) & "PERMS int NOT NULL DEFAULT 0)"
			
			SpecialSql13(SQL7) = "CREATE TABLE " & strTablePrefix & "ALLOWED_USERGROUPS ( "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "FORUM_ID int NOT NULL , "
			SpecialSql13(SQL7) = SpecialSql13(SQL7) & "PERMS int NOT NULL DEFAULT 0)"
			
			SpecialSql13(MySql) = "CREATE TABLE " & strTablePrefix & "ALLOWED_USERGROUPS ( "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "USERGROUP_ID int NOT NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "FORUM_ID int NOT NULL , "
			SpecialSql13(MySql) = SpecialSql13(MySql) & "PERMS int NOT NULL DEFAULT 0)"
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()
			
			
			'## Add the new config-values to the database
			strDummy = SetConfigValue(0,"STRSITEHEADER","0")
			strDummy = SetConfigValue(0,"STRSITELEFT","0")
			strDummy = SetConfigValue(0,"STRSITERIGHT","0")
			strDummy = SetConfigValue(0,"STRSITEFOOTER","0")
			strDummy = SetConfigValue(0,"STRSITEINTEGENABLED","0")
			strDummy = SetConfigValue(0,"STRSITEBORDER","0")
			strDummy = SetConfigValue(0,"STRUGVIEW","0")
			strDummy = SetConfigValue(0,"STRUGMEMVIEW","0")
			strDummy = SetConfigValue(0,"STRUGMODFORUMS","0")
			
			'Update code goes above this - delete this comment when done
			'## update the version info...
			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)

			UpdateErrorCode = UpdateErrorCheck()
			
			on error goto 0
			
			if UpdateErrorCode = 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
			elseif UpdateErrorCode = 2 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & vbNewLine)
			Response.Flush
		end if


'##########################################################################################################################
'##
'## end of update section, for newer versions add below this line (and then move this notice to the end of the new section)
'##
'##########################################################################################################################

		my_Conn.Close
		set my_Conn = nothing

		if intCriticalErrors = 0 then

			Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">The Upgrade has been completed without errors !</font></p>" & vbNewLine
		else
			Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4""><b>The Upgrade has NOT been completed without errors !</b></font></p>" & vbNewLine & _
					"<p><font face=""Verdana, Arial, Helvetica"" size=""2"">There were " & intCriticalErrors & "  Critical Errors...</font></p>" & vbNewLine
			if intWarnings > 0 then
				Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""2"">There were " & intWarnings & "  noncritical errors...</font></p>" & vbNewLine
			end if
		end if
		if intCriticalErrors > 0 then
			Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp?RC=3"">Click here to retry....</a></font></p>" & vbNewLine
		end if
		Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""default.asp"" target=""_top"">Click here to return to the forum....</a></font></p>" & vbNewLine & _
				"</center></div>" & vbNewLine

		Application(strCookieURL & "ConfigLoaded")= ""

	else
		Response.Redirect "setup_login.asp"
	end if

elseif ResponseCode = 5 then '## install a new database

	if strDBType = "access" then
		set mydbms_Conn = Server.CreateObject("ADODB.Connection")
		mydbms_Conn.Open strConnString
		strDBMSName = lcase(mydbms_Conn.Properties("DBMS Name"))
		mydbms_Conn.close
		set mydbms_Conn = nothing
	end if

	Response.Write	"<div align=""center""><center>" & vbNewLine & _
			"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">Installation of forum-tables in the database.</font></p>" & vbNewLine & _
			"</center></div>" & vbNewLine & _
			"<form action=""setup.asp?RC=6"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""left"">" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">Database Type:&nbsp;&nbsp;<b>"
	if strDBType = "access" then
		Response.Write("Microsoft Access")
		select case strDBMSName
			case "access"
				Response.Write(" (using ODBC Driver)")
			case "ms jet"
				Response.Write(" (using OLEDB Driver)")
		end select
	end if
	if strDBType = "sqlserver" then Response.Write("Microsoft SQL Server")
	if strDBType = "mysql" then Response.Write("MySQL")
	Response.Write	"</b></font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine
	if strDBType = "sqlserver" then
		Response.Write	"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""left""><p>" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Select the SQL-server version you are using:</b></font></p>" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2""><input type=""radio"" class=""radio"" name=""SQL_Server"" value=""SQL6"" >SQL-Server 6.5<br />" & vbNewLine & _
				"    <input type=""radio"" class=""radio"" checked name=""SQL_Server"" value=""SQL7"">SQL-Server 7 / 2000&nbsp;&nbsp;&nbsp;</p></font></p></td>" & vbNewLine & _
				"  </tr>" & vbNewLine
	end if
	if strDBType <> "access" then
		Response.Write	"  <tr>" & vbNewLine & _
				"    <td bgColor=""#9FAFDF"" align=""left"">" & vbNewLine & _
				"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">To install the tables in the database you need to create the empty database on the server first.  " & vbNewLine & _
				"    Then you have to provide a username and password of a user that has table creation/modification rights at the database you use.  " & vbNewLine & _
				"    This might not be the same user as you use in your connectionstring !</font></p>" & vbNewLine & _
				"      <table border=""0"" align=""left"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Name:</b>&nbsp;</font></td>" & vbNewLine & _
				"          <td align=""left""><input type=""text"" name=""DBUserName"" size=""25"" style=""width:150px;""></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Password:</b>&nbsp;</font></td>" & vbNewLine & _
				"          <td align=""left""><input type=""password"" name=""DBPassword"" size=""25"" style=""width:150px;""></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine
	end if
	Response.Write	"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""left"">" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Forum Admin UserName/Password:</b></font></p>" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">Here you will choose the Forum Admin UserName & Password that will be entered into the database for the Forum Admin.  " & vbNewLine & _
			"    The password should be something that you can remember, but not something easily guessed by anyone else.  Size limit is 25 characters.</font></p>" & vbNewLine & _
			"      <table border=""0"" align=""left"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Username:</b>&nbsp;</font></td>" & vbNewLine & _
			"          <td align=""left""><input maxLength=""25"" type=""text"" name=""AdminName"" size=""25"" style=""width:150px;""></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Password:</b>&nbsp;</font></td>" & vbNewLine & _
			"          <td align=""left""><input maxLength=""25"" type=""password"" name=""AdminPassword"" size=""25"" style=""width:150px;""></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Password Again:</b>&nbsp;</font></td>" & vbNewLine & _
			"          <td align=""left""><input maxLength=""25"" type=""password"" name=""AdminPassword2"" size=""25"" style=""width:150px;""></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"    </td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center""><input type=""submit"" value=""Continue"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</form>" & vbNewLine

elseif ResponseCode = 6 then '## start installing the tables in the database

	Err_Msg = ""

	strAdminName = trim(chkString(Request.Form("AdminName"),"SQLString"))
	strAdminPassword = trim(chkString(Request.Form("AdminPassword"),"SQLString"))
	strAdminPassword2 = trim(chkString(Request.Form("AdminPassword2"),"SQLString"))

	if strAdminName = "" then
		Err_Msg = Err_Msg & "<li>You must choose the Forum Admin UserName</li>"
	end if
	
	if Len(strAdminName) < 3 then
		Err_Msg = Err_Msg & "<li>Your Forum Admin UserName must be at least <strong>3</strong> characters long</li>"
	end if

	if not IsValidString(strAdminName) then
		Err_Msg = Err_Msg & "<li>You may not use any of these chars in the Forum Admin UserName  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
	end if

	if strAdminPassword = "" then
		Err_Msg = Err_Msg &  "<li>You must choose the Forum Admin Password</li>"
	end if

	if strAdminPassword <> strAdminPassword2 then
		Err_Msg = Err_Msg & "<li>Your Forum Admin Passwords didn't match</li>"
	end if

	if not IsValidString(strAdminPassword) then
		Err_Msg = Err_Msg & "<li>You may not use any of these chars in the Forum Admin Password  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
	end if

	if Err_Msg <> "" then
		Response.Write	"<table border=""0"" height=""50%"" align=""center"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine & _
				"    <p align=""center""><font face=""Verdana, Arial, Helvetica"" size=""4"" color=""#FF0000"">There has been a problem!</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""Verdana, Arial, Helvetica"" size=""2"" color=""#FF0000""><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"    <p align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""JavaScript:history.go(-1)"">Go back to correct the problem</a></font></p>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine
		Response.End
	end if

	strAdminPassword = sha256("" & strAdminPassword)

	on error resume next

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

	if strDBType = "access" then
		strDBMSName = lcase(my_Conn.Properties("DBMS Name"))
	end if

	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDesc = my_conn.Errors(counter).Description
		if ConnErrorNumber <> 0 then
			my_Conn.Errors.Clear
			Err.Clear
			Response.Redirect "setup.asp?RC=1&CC=1&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc)
		end if
	next

	my_Conn.Errors.Clear
	Err.Clear

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_LEVEL = 3"

	Set rs = my_Conn.Execute(strSql)

	blnError = FALSE
	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		if ConnErrorNumber <> 0 then
			my_Conn.Errors.Clear
			Err.Clear
			blnError = TRUE
		end if
	next

	If not(blnError) then
		if (not(rs.BOF or rs.EOF) and (Session(strCookieURL & "Approval") <> strAdminCode) ) then
			if strDBType = "access" then
				Response.Write	"<div align=""center""><center>" & vbNewLine & _
						"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">The Forum Tables have already been installed.</font></p>" & vbNewLine & _
						"<p><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""default.asp"">Click here to go to the Forum</a></font></p>" & vbNewLine & _
						"</center></div>" & vbNewLine
				Response.end
			end if
			Response.Write	"<div align=""center""><center>" & vbNewLine & _
					"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">You need to logon first.</font></p>" & vbNewLine & _
					"</center></div>" & vbNewLine & _
					"<form action=""setup_login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
					"<input type=""hidden"" name=""setup"" value=""Y"">" & vbNewLine & _
					"<input type=""hidden"" name=""ReturnTo"" value=""RC=5"">" & vbNewLine & _
					"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
					"  <tr>" & vbNewLine & _
					"    <td bgColor=""#9FAFDF"" align=""left"">" & vbNewLine & _
					"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
					"    To Re-install the tables you need to be logged on as a forum administrator.<br />" & vbNewLine
			if strSender <> "" then
				Response.Write	"    If you are not the Administrator of this forum<br /> please report this error here: <a href=""mailto:" & strSender & """>" & strSender & "</a>.<br /><br />" & vbNewLine
			end if
			Response.Write	"    </font></p></td>" & vbNewLine & _
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
					"  </tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
			Response.end
		end if
	end if

	rs.close
	Set rs = nothing

	my_Conn.Errors.Clear
	Err.Clear

	on error goto 0

	Response.Write	"<div align=""center""><center>" & vbNewLine & _
			"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">Please Wait until the installation has been completed !</font></p>" & vbNewLine

	if strDBType = "access" or not Instr(strConnString,"uid=") > 0 then
		strInstallString = strConnString
	else
		strInstallString = CreateConnectionString(strConnString, Request.Form("DBUserName"), Request.Form("DBPassword"))
	end if

	on error resume next

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strInstallString

	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDesc = my_conn.Errors(counter).Description
		if ConnErrorNumber <> 0 then 
			my_Conn.Errors.Clear
			Err.Clear 
			Response.Redirect "setup.asp?RC=1&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc) & "&RET=" & Server.URLEncode("setup.asp?RC=5")
		end if
	next

	on error goto 0

	intCriticalErrors = 0

	Response.Write	"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine
	strSQL_Server = Request.Form("Sql_Server")
	if strDBType = "mysql" then
%>
		<!--#INCLUDE FILE="inc_create_forum_mysql.asp" -->
<%
	elseif strDBType = "sqlserver" then
		if strSQL_Server = "SQL6" then
			strN = ""
		else
			strN = "n"
		end if
%>
		<!--#INCLUDE FILE="inc_create_forum_mssql.asp" -->
<%
	elseif strDBType = "access" then
		if strDBMSName = "ms jet" then
			strN = "n"
		else
			strN = ""
		end if
%>
		<!--#INCLUDE FILE="inc_create_forum_access.asp" -->
<%
	end if
	Response.Write	"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine

	my_Conn.Close
	set my_Conn = nothing

	if intCriticalErrors = 0 then
		Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4"">The Installation has been completed !</font></p>" & vbNewLine
	else
		Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""4""><b>The Installation has NOT been completed !</b></font></p>" & vbNewLine & _
				"<p><font face=""Verdana, Arial, Helvetica"" size=""2"">There were " & intCriticalErrors & "  Critical Errors...</font></p>" & vbNewLine
	end if
	if intCriticalErrors > 0 then
		Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp?RC=5"">Click here to retry....</a></font></p>" & vbNewLine
	end if
	Response.Write	"<p><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp"" target=""_top"">Click here to check the Database....</a></font></p>" & vbNewLine & _
			"</center></div>" & vbNewLine
else
	Response.Write	"<html>" & vbNewLine & _
			vbNewLine & _
			"<head>" & vbNewLine & _
			"<title>Forum-Setup Page</title>" & vbNewLine

	'## START - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
	'## END   - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

	Response.Write	"<style><!--" & vbNewLine & _
			"a:link    {color:darkblue;text-decoration:underline}" & vbNewLine & _
			"a:visited {color:blue;text-decoration:underline}" & vbNewLine & _
			"a:hover   {color:red;text-decoration:underline}" & vbNewLine & _
			"--></style>" & vbNewLine & _
			"</head>" & vbNewLine & _
			vbNewLine & _
			"<body bgColor=""white"" text=""midnightblue"" link=""darkblue"" aLink=""red"" vLink=""red"" onLoad=""window.focus()"">" & vbNewLine & _
			"<div align=""center""><center>" & vbNewLine & _
			"<font face=""Verdana, Arial, Helvetica"" size=""4"">There has been an error !!</font>" & vbNewLine & _
			"</center></div>" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" align=""left"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
			"    <font face=""Verdana, Arial, Helvetica"" size=""2"">"
	Response.Write(HEX(Err.number) & ", " & Err.description)
	Response.Write	"    </font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center"">" & vbNewLine & _
			"    <font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    <a href=""default.asp"" target=""_top"">Click here to retry.</a>" & vbNewLine & _
			"    </font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
end if
Response.Write	"</body>" & vbNewLine & _
		vbNewLine & _
		"</html>" & vbNewLine

sub CheckSqlError()

	dim ChkConnErrorNumber

	for counter = 0 to my_Conn.Errors.Count -1
		ChkConnErrorNumber = Err.Number
		if ChkConnErrorNumber <> 0 then

			my_Conn.Errors.Clear
			Err.Clear

			strSql = "SELECT " & strTablePrefix & "CONFIG.C_STRVERSION, "
			strSql = strSql & strTablePrefix & "CONFIG.C_STRSENDER "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG "

			set rsInfo = my_Conn.Execute (StrSql)
			strVersion = rsInfo("C_STRVERSION")
			strSender = rsInfo("C_STRSENDER")

			rsInfo.Close
			set rsInfo = nothing
			my_Conn.Close
			set my_Conn = nothing

			Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & ChkConnErrorNumber
		end if
	next
end sub

sub CheckSqlErrorNew()

	dim ChkConnErrorNumber

	for counter = 0 to my_Conn.Errors.Count -1
		ChkConnErrorNumber = Err.Number
		if ChkConnErrorNumber <> 0 then

			my_Conn.Errors.Clear
			Err.Clear

			strSql = "SELECT C_VALUE "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			set rsInfo = my_Conn.Execute (StrSql)
			strVersion = rsInfo("C_VALUE")

			strSql = "SELECT C_VALUE "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " WHERE C_VARIABLE = 'STRSENDER'"

			set rsInfo = my_Conn.Execute (StrSql)
			strSender = rsInfo("C_VALUE")

			rsInfo.Close
			set rsInfo = nothing
			my_Conn.Close
			set my_Conn = nothing

			Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & ChkConnErrorNumber
		end if
	next
end sub

function UpdateErrorCheck()

	dim intErrorNumber
	dim counter

	intErrorNumber = 0
	for counter = 0 to my_Conn.Errors.Count -1
		intErrorNumber = my_Conn.Errors(counter).Number
		if intErrorNumber <> 0 or Err.Number <> 0 then  
			select case intErrorNumber
				case -2147217900, -2147217887
					UpdateErrorCheck = 1
					counter = my_Conn.Errors.Count -1
				case -2147467259
					UpdateErrorCheck = 2
					if strDBType = "mysql" then
						if instr(my_Conn.Errors(counter).Description, "Duplicate column name") > 0 then
							UpdateErrorCheck = 1
						end if
					end if
					counter = my_Conn.Errors.Count -1	
				case else
					UpdateErrorCheck = intErrorNumber
			end select
		end if
	next
end function

Sub AddColumns(Columns, intCriticalErrors, intWarnings)

	Dim colCounter

	Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine)
	For colCounter = 0 to Ubound(Columns, 1) 
		on error resume next
		my_Conn.Errors.Clear
		Err.Clear
		
		strUpdateSql = "ALTER TABLE " & Columns(colCounter, Prefix) & Columns(colCounter, TableName) & "  "
		if strDBType = "access" then
			strUpdateSql = strUpdateSql & " ADD COLUMN " & Columns(colCounter, FieldName) & " " 
		else
			strUpdateSql = strUpdateSql & " ADD " & Columns(colCounter, FieldName) & " " 
		end if
		if strDBType = "access" then
			strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_Access) & " " & Columns(colCounter, ConstraintAccess) & " "
		elseif strDBType = "sqlserver" then 
			if strSQL_Server = "SQL7" then
				strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_SQL7) & " " & Columns(colCounter, ConstraintSQL7) & " "
			else
				strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_SQL6) & " " & Columns(colCounter, ConstraintSQL6) & " "
			end if
		elseif strDBType = "mysql" then
			strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_MySql) & " " & Columns(colCounter, ConstraintMySql) & " "
		end if

		my_Conn.Execute strUpdateSql

		UpdateErrorCode = UpdateErrorCheck()

		on error goto 0

		if UpdateErrorCode = 0 then
			Response.Write("  <tr>" & vbNewLine)
			Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded:</b></font></td>" & vbNewLine)
			Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Columns(colCounter, FieldName) & " has been added to the " & Columns(colCounter, TableName) & " table</font></td>" & vbNewLine)
			Response.Write("  </tr>" & vbNewLine)
		elseif UpdateErrorCode = 1 then
			Response.Write("  <tr>" & vbNewLine)
			Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Noncritical error: </b></font></td>" & vbNewLine)
			Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Columns(colCounter, Fieldname) & " already existed in the " & Columns(colCounter, TableName) & " table</font></td>" & vbNewLine)
			Response.Write("  </tr>" & vbNewLine)
			intWarnings = intWarnings + 1
		elseif UpdateErrorCode = 2 then
			Response.Write("  <tr>" & vbNewLine)
			Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td>" & vbNewLine)
			Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> No write access to the table " & Columns(colCounter, TableName) & " <br />" & Columns(colCounter, Fieldname) & " not added to database!</font></td>" & vbNewLine)
			Response.Write("  </tr>" & vbNewLine)
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("  <tr>" & vbNewLine)
			Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td>" & vbNewLine)
			Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " in statement [" & strUpdateSql & "] while trying to add " & Columns(colCounter, Fieldname) & " to the " & Columns(colCounter, TableName) & " table</font></td>" & vbNewLine)
			Response.Write("  </tr>" & vbNewLine)
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Flush
	Next
	Response.Write("</table>" & vbNewLine)
end sub

sub SpecialUpdates(strSql, strOkMessage)

	dim strUpdateSql
	dim SpecialErrors

	on error resume next
	my_Conn.Errors.Clear
	Err.Clear
		
	if strDBType = "access" then
		strUpdateSql = strSql(Access) 
	elseif strDBType = "sqlserver" then 
		if strSQL_Server = "SQL7" then
			strUpdateSql = strSql(SQL7)
		else
			strUpdateSql = strSql(SQL6)
		end if
	elseif strDBType = "mysql" then
		strUpdateSql = strSql(MySql)
	end if

	my_Conn.Execute strUpdateSql
	
	SpecialErrors = 0
	Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" align=""center"">" & vbNewLine)
	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDescription = my_Conn.Errors(counter).Description

		if ConnErrorNumber <> 0 then 

			if ConnErrorNumber = -2147217900 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			elseif (instr(1,my_Conn.Errors(counter).Description,"Table",1) > 0) and (instr(1,my_Conn.Errors(counter).Description,"does not exist",1) > 0) and (instr(1,strUpdateSql,"DROP TABLE",1) > 0) then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			elseif strDBType = "mysql" and instr(my_Conn.Errors(counter).Description, "already exists") > 0 then
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""orange"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			else
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				Response.Write("  <tr>" & vbNewLine)
				Response.Write("    <td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></font></td>" & vbNewLine)
				Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</font></td>" & vbNewLine)
				Response.Write("  </tr>" & vbNewLine)
				intCriticalErrors = intCriticalErrors + 1
				SpecialErrors = 1
			end if
		end if
	next
	if SpecialErrors = 0 then
		Response.Write("  <tr>" & vbNewLine)
		Response.Write("    <td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>" & vbNewLine)
		Response.Write("    <td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & strOkMessage & "</font></td>" & vbNewLine)
		Response.Write("  </tr>" & vbNewLine)
	end if

	Response.Write("</table>" & vbNewLine)
	Response.Flush

	my_Conn.Errors.Clear 
	Err.Clear
	on error goto 0
end sub

function CreateConnectionString(strConn, UserName, Password)
	'## strConnString = "driver={SQL Server};server=YYYY;uid=XX;pwd=XXXX;database=ZZZZZZ"
	'## strConnString = "driver={SQL Server};server=YYYY;USER ID=XX;PASSWORD=XXXX;database=ZZZZZZ"

	Dim TempConnString
	Dim uidTagStart(1), uidTagEnd
	Dim pwdTagStart(1), pwdTagEnd
	Dim uidTagStartPos, uidTagEndPos
	Dim pwdTagStartPos, pwdTagEndPos
	Dim blnUIDok, blnPWDok

	uidTagStart(0)		= "UID="
	uidTagStart(1)		= "USER ID="
	uidTagEnd		= ";" 
	pwdTagStart(0)		= "PWD="
	pwdTagStart(1)		= "PASSWORD="
	pwdTagEnd		= ";"

	TempConnString = strConn
	blnUIDok = FALSE
	blnPWDok = FALSE

	for Counter = 0 to Ubound(uidTagStart)

		uidTagStartPos = InStr(1, UCase(TempConnString), UCase(uidTagStart(Counter)), 1)
		if uidTagStartPos > 0 then 
			uidTagEndPos = InStr(uidTagStartPos, UCase(TempConnString), UCase(uidTagEnd), 1)
		else
			uidTagEndPos = 0
		end if

		if (uidTagStartpos > 0) and (uidTagEndPos > 0) then
			TempConnString = Left(TempConnString, (uidTagStartPos + len(uidTagStart(Counter))-1)) & UserName & Right(TempConnString, (len(TempConnString) - uidTagEndPos) + 1)
			blnUIDok = TRUE
		end if

		pwdTagStartPos = InStr(1, TempConnString, pwdTagStart(Counter), 1)
		if pwdTagStartPos > 0 then
			pwdTagEndPos = InStr(pwdTagStartPos, TempConnString, pwdTagEnd, 1)
		else
			pwdTagEndPos = 0
		end if

		if (pwdTagStartpos > 0) and (pwdTagEndPos > 0) then
			TempConnString = Left(TempConnString, (pwdTagStartPos + len(pwdTagStart(Counter))-1)) & Password & Right(TempConnString, (len(TempConnString) - pwdTagEndPos) + 1)
			blnPWDok = TRUE
		end if	

	next
	if blnUIDok and blnPWDok then
		CreateConnectionString = TempConnString
	else
		CreateConnectionString = "<Error>"
	end if
end function

Sub TransferOldConfig

	on error resume next

	strSql = "SELECT * FROM " & strTablePrefix & "CONFIG WHERE CONFIG_ID = " & 1

	set rs = my_Conn.Execute(strSql)

	Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")

	UpdateErrorCode = UpdateErrorCheck()

	if UpdateErrorCode = 0 then
		if not(rs.eof) then 
			strDummy = SetConfigValue(0,"STRVERSION", rs("C_STRVERSION"))
			strDummy = SetConfigValue(0,"STRFORUMTITLE", rs("C_STRFORUMTITLE"))
			strDummy = SetConfigValue(0,"STRCOPYRIGHT" , rs("C_STRCOPYRIGHT"))
			strDummy = SetConfigValue(0,"STRTITLEIMAGE", rs("C_STRTITLEIMAGE"))
			strDummy = SetConfigValue(0,"STRHOMEURL", rs("C_STRHOMEURL"))
			strDummy = SetConfigValue(0,"STRFORUMURL", rs("C_STRFORUMURL"))
			strDummy = SetConfigValue(0,"STRAUTHTYPE", rs("C_STRAUTHTYPE"))
			strDummy = SetConfigValue(0,"STRSETCOOKIETOFORUM", rs("C_STRSETCOOKIETOFORUM"))
			strDummy = SetConfigValue(0,"STREMAIL", rs("C_STREMAIL"))
			strDummy = SetConfigValue(0,"STRUNIQUEEMAIL", rs("C_STRUNIQUEEMAIL"))
			strDummy = SetConfigValue(0,"STRMAILMODE", rs("C_STRMAILMODE"))
			strDummy = SetConfigValue(0,"STRMAILSERVER", rs("C_STRMAILSERVER"))
			strDummy = SetConfigValue(0,"STRSENDER", rs("C_STRSENDER"))
			strDummy = SetConfigValue(0,"STRDATETYPE", rs("C_STRDATETYPE"))
			strDummy = SetConfigValue(0,"STRTIMETYPE", rs("C_STRTIMETYPE"))
			strDummy = SetConfigValue(0,"STRTIMEADJUSTLOCATION", rs("C_STRTIMEADJUSTLOCATION"))
			strDummy = SetConfigValue(0,"STRTIMEADJUST", rs("C_STRTIMEADJUST"))
			strDummy = SetConfigValue(0,"STRMOVETOPICMODE", rs("C_STRMOVETOPICMODE"))
			strDummy = SetConfigValue(0,"STRPRIVATEFORUMS", rs("C_STRPRIVATEFORUMS"))
			strDummy = SetConfigValue(0,"STRSHOWMODERATORS", rs("C_STRSHOWMODERATORS"))
			strDummy = SetConfigValue(0,"STRSHOWRANK", rs("C_STRSHOWRANK"))
			strDummy = SetConfigValue(0,"STRHIDEEMAIL", rs("C_STRHIDEEMAIL"))
			strDummy = SetConfigValue(0,"STRIPLOGGING", rs("C_STRIPLOGGING"))
			strDummy = SetConfigValue(0,"STRALLOWFORUMCODE", rs("C_STRALLOWFORUMCODE"))
			strDummy = SetConfigValue(0,"STRIMGINPOSTS", rs("C_STRIMGINPOSTS") )
			strDummy = SetConfigValue(0,"STRALLOWHTML", rs("C_STRALLOWHTML"))
			strDummy = SetConfigValue(0,"STRSECUREADMIN", rs("C_STRSECUREADMIN"))
			strDummy = SetConfigValue(0,"STRNOCOOKIES", rs("C_STRNOCOOKIES"))
			strDummy = SetConfigValue(0,"STREDITEDBYDATE", rs("C_STREDITEDBYDATE"))
			strDummy = SetConfigValue(0,"STRHOTTOPIC", rs("C_STRHOTTOPIC"))
			strDummy = SetConfigValue(0,"INTHOTTOPICNUM", rs("C_INTHOTTOPICNUM"))
			strDummy = SetConfigValue(0,"STRHOMEPAGE", rs("C_STRHOMEPAGE"))
			strDummy = SetConfigValue(0,"STRAIM", rs("C_STRAIM"))
			strDummy = SetConfigValue(0,"STRYAHOO", rs("C_STRYAHOO"))
			strDummy = SetConfigValue(0,"STRICQ", rs("C_STRICQ"))
			strDummy = SetConfigValue(0,"STRICONS", rs("C_STRICONS"))
			strDummy = SetConfigValue(0,"STRGFXBUTTONS", rs("C_STRGFXBUTTONS"))
			strDummy = SetConfigValue(0,"STRBADWORDFILTER", rs("C_STRBADWORDFILTER"))
			strDummy = SetConfigValue(0,"STRBADWORDS", rs("C_STRBADWORDS"))
			strDummy = SetConfigValue(0,"STRDEFAULTFONTFACE", rs("C_STRDEFAULTFONTFACE"))
			strDummy = SetConfigValue(0,"STRDEFAULTFONTSIZE ", rs("C_STRDEFAULTFONTSIZE"))
			strDummy = SetConfigValue(0,"STRHEADERFONTSIZE", rs("C_STRHEADERFONTSIZE"))
			strDummy = SetConfigValue(0,"STRFOOTERFONTSIZE", rs("C_STRFOOTERFONTSIZE"))
			strDummy = SetConfigValue(0,"STRPAGEBGCOLOR", rs("C_STRPAGEBGCOLOR"))
			strDummy = SetConfigValue(0,"STRDEFAULTFONTCOLOR", rs("C_STRDEFAULTFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRLINKCOLOR", rs("C_STRLINKCOLOR"))
			strDummy = SetConfigValue(0,"STRLINKTEXTDECORATION", rs("C_STRLINKTEXTDECORATION"))
			strDummy = SetConfigValue(0,"STRVISITEDLINKCOLOR", rs("C_STRVISITEDLINKCOLOR"))
			strDummy = SetConfigValue(0,"STRVISITEDTEXTDECORATION", rs("C_STRVISITEDTEXTDECORATION"))
			strDummy = SetConfigValue(0,"STRACTIVELINKCOLOR", rs("C_STRACTIVELINKCOLOR"))
			strDummy = SetConfigValue(0,"STRHOVERFONTCOLOR", rs("C_STRHOVERFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRHOVERTEXTDECORATION", rs("C_STRHOVERTEXTDECORATION"))
			strDummy = SetConfigValue(0,"STRHEADCELLCOLOR", rs("C_STRHEADCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRHEADFONTCOLOR", rs("C_STRHEADFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRCATEGORYCELLCOLOR", rs("C_STRCATEGORYCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRCATEGORYFONTCOLOR", rs("C_STRCATEGORYFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMFIRSTCELLCOLOR", rs("C_STRFORUMFIRSTCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMCELLCOLOR", rs("C_STRFORUMCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRALTFORUMCELLCOLOR", rs("C_STRALTFORUMCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMFONTCOLOR", rs("C_STRFORUMFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMLINKCOLOR", rs("C_STRFORUMLINKCOLOR"))
			strDummy = SetConfigValue(0,"STRTABLEBORDERCOLOR", rs("C_STRTABLEBORDERCOLOR"))
			strDummy = SetConfigValue(0,"STRPOPUPTABLECOLOR", rs("C_STRPOPUPTABLECOLOR"))
			strDummy = SetConfigValue(0,"STRPOPUPBORDERCOLOR", rs("C_STRPOPUPBORDERCOLOR"))
			strDummy = SetConfigValue(0,"STRNEWFONTCOLOR", rs("C_STRNEWFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRTOPICWIDTHLEFT", rs("C_STRTOPICWIDTHLEFT"))
			strDummy = SetConfigValue(0,"STRTOPICWIDTHRIGHT", rs("C_STRTOPICWIDTHRIGHT"))
			strDummy = SetConfigValue(0,"STRTOPICNOWRAPLEFT", rs("C_STRTOPICNOWRAPLEFT"))
			strDummy = SetConfigValue(0,"STRTOPICNOWRAPRIGHT", rs("C_STRTOPICNOWRAPRIGHT"))
			strDummy = SetConfigValue(0,"STRRANKADMIN", rs("C_STRRANKADMIN"))
			strDummy = SetConfigValue(0,"STRRANKMOD", rs("C_STRRANKMOD"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL0", rs("C_STRRANKLEVEL0"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL1", rs("C_STRRANKLEVEL1"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL2", rs("C_STRRANKLEVEL2"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL3", rs("C_STRRANKLEVEL3"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL4", rs("C_STRRANKLEVEL4"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL5", rs("C_STRRANKLEVEL5"))
			strDummy = SetConfigValue(0,"STRRANKCOLORADMIN", rs("C_STRRANKCOLORADMIN"))
			strDummy = SetConfigValue(0,"STRRANKCOLORMOD", rs("C_STRRANKCOLORMOD"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR0", rs("C_STRRANKCOLOR0"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR1", rs("C_STRRANKCOLOR1"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR2", rs("C_STRRANKCOLOR2"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR3", rs("C_STRRANKCOLOR3"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR4", rs("C_STRRANKCOLOR4"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR5", rs("C_STRRANKCOLOR5"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL0", rs("C_INTRANKLEVEL0"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL1", rs("C_INTRANKLEVEL1"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL2", rs("C_INTRANKLEVEL2"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL3", rs("C_INTRANKLEVEL3"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL4", rs("C_INTRANKLEVEL4"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL5", rs("C_INTRANKLEVEL5"))
			strDummy = SetConfigValue(0,"STRSIGNATURES", rs("C_STRSIGNATURES") )
			strDummy = SetConfigValue(0,"STRSHOWSTATISTICS", rs("C_STRSHOWSTATISTICS"))
			strDummy = SetConfigValue(0,"STRSHOWIMAGEPOWEREDBY", rs("C_STRSHOWIMAGEPOWEREDBY"))
			strDummy = SetConfigValue(0,"STRLOGONFORMAIL", rs("C_STRLOGONFORMAIL"))
			strDummy = SetConfigValue(0,"STRSHOWPAGING", rs("C_STRSHOWPAGING"))
			strDummy = SetConfigValue(0,"STRSHOWTOPICNAV", rs("C_STRSHOWTOPICNAV"))
			strDummy = SetConfigValue(0,"STRPAGESIZE", rs("C_STRPAGESIZE"))
			strDummy = SetConfigValue(0,"STRPAGENUMBERSIZE", rs("C_STRPAGENUMBERSIZE"))
			strDummy = SetConfigValue(0,"STRFULLNAME", rs("C_STRFULLNAME"))
			strDummy = SetConfigValue(0,"STRPICTURE", rs("C_STRPICTURE"))
			strDummy = SetConfigValue(0,"STRSEX", rs("C_STRSEX"))
			strDummy = SetConfigValue(0,"STRCITY", rs("C_STRCITY"))
			strDummy = SetConfigValue(0,"STRSTATE", rs("C_STRSTATE"))
			strDummy = SetConfigValue(0,"STRAGE", rs("C_STRAGE"))
			strDummy = SetConfigValue(0,"STRCOUNTRY", rs("C_STRCOUNTRY"))
			strDummy = SetConfigValue(0,"STROCCUPATION", rs("C_STROCCUPATION"))
			strDummy = SetConfigValue(0,"STRHOMEPAGE", rs("C_STRHOMEPAGE"))
			strDummy = SetConfigValue(0,"STRFAVLINKS", rs("C_STRFAVLINKS"))
			strDummy = SetConfigValue(0,"STRBIO", rs("C_STRBIO"))
			strDummy = SetConfigValue(0,"STRHOBBIES", rs("C_STRHOBBIES"))
			strDummy = SetConfigValue(0,"STRLNEWS", rs("C_STRLNEWS"))
			strDummy = SetConfigValue(0,"STRQUOTE", rs("C_STRQUOTE"))
			strDummy = SetConfigValue(0,"STRMARSTATUS", rs("C_STRMARSTATUS"))
			strDummy = SetConfigValue(0,"STRRECENTTOPICS", rs("C_STRRECENTTOPICS"))
			strDummy = SetConfigValue(0,"STRNTGROUPS", rs("C_STRNTGROUPS"))
			strDummy = SetConfigValue(0,"STRAUTOLOGON", rs("C_STRAUTOLOGON"))
			strDummy = SetConfigValue(0,"STRMOVENOTIFY", "1")
			strDummy = SetConfigValue(0,"STRSUBSCRIPTION", "1")
			strDummy = SetConfigValue(0,"STRMODERATION", "1")

			Response.Write("<tr><td bgColor=""green"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> Config values transferred to new table</font></td></tr>")
		else
			Response.Write("<tr><td bgColor=orange align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2""> No existing config values found</font></td></tr>")
		end if
	else
		Response.Write("<tr><td bgColor=""red"" align=""left"" width=""30%""><font face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=""#9FAFDF"" align=""left""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to tranfer the existing config values to the new table</font></td></tr>")
		intCriticalErrors = intCriticalErrors + 1
	end if
	Response.Write("</table>")
	
	rs.close
	set rs = nothing

	on error goto 0
end sub

Sub UpDateAccessFields(pOldversion)

	if pOldversion <= 5  then

		my_Conn.execute ("UPDATE " & strTablePrefix & "CATEGORY SET CAT_MODERATION = 0 WHERE (CAT_MODERATION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "CATEGORY SET CAT_SUBSCRIPTION = 0 WHERE (CAT_SUBSCRIPTION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "CATEGORY SET CAT_ORDER = 1 WHERE (CAT_ORDER Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_L_ARCHIVE = '' WHERE (F_L_ARCHIVE Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_ARCHIVE_SCHED = 30 WHERE (F_ARCHIVE_SCHED Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_L_DELETE = '' WHERE (F_L_DELETE Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_DELETE_SCHED = 365 WHERE (F_DELETE_SCHED Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_MODERATION = 0 WHERE (F_MODERATION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_SUBSCRIPTION = 0 WHERE (F_SUBSCRIPTION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_ORDER = 1 WHERE (F_ORDER Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "TOPICS SET T_ARCHIVE_FLAG = 1 WHERE (T_ARCHIVE_FLAG Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "REPLY SET R_STATUS = 0 WHERE (R_STATUS Is Null)")

		on error resume next
		my_Conn.execute("ALTER TABLE " & strTablePrefix & "TOPICS DROP COLUMN C_STRMOVENOTIFY")
		on error goto 0
	end if
end sub

function SetConfigValue(bUpdate, fVariable, fValue)

	' bUpdate = 1 : if it exists then overwrite with new values
	' bUpdate = 0 : if it exists then leave unchanged

	Dim strSql

	strSql = "SELECT C_VARIABLE FROM " & strTablePrefix & "CONFIG_NEW " &_
		 " WHERE C_VARIABLE = '" & fVariable & "' "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if (rs.EOF or rs.BOF) then '## New config-value
		SetConfigValue = "added"
		my_conn.execute ("INSERT INTO " & strTablePrefix & "CONFIG_NEW (C_VALUE,C_VARIABLE) VALUES ('" & fValue & "' , '" & fVariable & "')")
	else
		if bUpdate <> 0 then 
			SetConfigValue = "updated"
			my_conn.execute ("UPDATE " & strTablePrefix & "CONFIG_NEW SET C_VALUE = '" & fValue & "' WHERE C_VARIABLE = '" & fVariable &"'")
		else ' not changed
			SetConfigValue = "unchanged"
		end if
	end if

	rs.close
	set rs = nothing
end function

function SetBadWordValue(bUpdate, fVariable, fValue)

	' bUpdate = 1 : if it exists then overwrite with new values
	' bUpdate = 0 : if it exists then leave unchanged

	Dim strSql

	strSql = "SELECT B_BADWORD FROM " & strFilterTablePrefix & "BADWORDS " &_
		 " WHERE B_BADWORD = '" & fVariable & "' "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if (rs.EOF or rs.BOF) then '## New Badword
		SetBadWordValue = "added"
		my_conn.execute ("INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_REPLACE,B_BADWORD) VALUES ('" & chkString(fValue,"sqlstring") & "' , '" & fVariable & "')")
	else
		if bUpdate <> 0 then 
			SetBadWordValue = "updated"
			my_conn.execute ("UPDATE " & strFilterTablePrefix & "BADWORDS SET B_REPLACE = '" & chkString(fValue,"sqlstring") & "' WHERE B_BADWORD = '" & fVariable &"'")
		else ' not changed
			SetBadWordValue = "unchanged"
		end if
	end if

	rs.close
	set rs = nothing
end function

function doublenum(fNum)
	if fNum > 9 then 
		doublenum = fNum 
	else
		doublenum = "0" & fNum
	end if
end function

function DateToStr(dtDateTime)
	if not isDate(dtDateTime) then
		dtDateTime = strToDate(dtDateTime)
	end if

	DateToStr = year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime)) & ""
end function

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

function CheckSelected(ByVal chkval1, chkval2)
	if IsNumeric(chkval1) then chkval1 = cLng(chkval1)
	if (chkval1 = chkval2) then
		CheckSelected = " selected"
	else
		CheckSelected = ""
	end if
end function

function HTMLEncode(pString)
	fString = trim(pString)
	if fString = "" or IsNull(fString) then fString = " "
	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")
	HTMLEncode = fString
end function

function chkString(pString,fField_Type) '## Types - SQLString
	fString = trim(pString)
	if fString = "" or isNull(fString) then
		fString = " "
	end if
	select case fField_Type
		case "SQLString"
			fString = Replace(fString, "'", "''")
			if strDBType = "mysql" then
				fString = Replace(fString, "\0", "\\0")
				fString = Replace(fString, "\'", "\\'")
				fString = Replace(fString, "\""", "\\""")
				fString = Replace(fString, "\b", "\\b")
				fString = Replace(fString, "\n", "\\n")
				fString = Replace(fString, "\r", "\\r")
				fString = Replace(fString, "\t", "\\t")
				fString = Replace(fString, "\z", "\\z")
				fString = Replace(fString, "\%", "\\%")
				fString = Replace(fString, "\_", "\\_")
			end if
			fString = HTMLEncode(fString)
			chkString = fString
			exit function
	end select
	chkString = fString
end function
%>
