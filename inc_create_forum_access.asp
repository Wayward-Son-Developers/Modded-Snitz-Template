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

if (strDBType = "access") then
	my_Conn.Errors.Clear
	on error resume next

	strSql = "CREATE TABLE " & strTablePrefix & "CONFIG_NEW ( "
	if strDBMSName = "ms jet" then
		strSql = strSql & "ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	else
		strSql = strSql & "ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
	end if
	strSql = strSql & "C_VARIABLE " & strN & "varchar (255) NULL , "
	strSql = strSql & "C_VALUE " & strN & "varchar (255) NULL )"

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	'ChkDBInstall()

	'#####################################################
	'## Insert Default Configuration Data Into Database ##
	'#####################################################
%>
	<!--#INCLUDE FILE="inc_create_forum_configvalues.asp"-->
<%
	'#######################################
	'## Insert Default Data Into Database ##
	'#######################################

	strSql = "SELECT CAT_ID FROM " & strTablePrefix & "CATEGORY"
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "CATEGORY(CAT_STATUS, CAT_NAME) VALUES(1, 'Snitz Forums 2000')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS WHERE M_LEVEL = 3"
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERS (M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, "
		strSql = strSql & "M_HOMEPAGE, M_LINK1, M_LINK2, M_PHOTO_URL, M_SIG, M_VIEW_SIG, M_SIG_DEFAULT, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_ICQ, M_MSN, M_YAHOO, "
		strSql = strSql & "M_POSTS, M_DATE, M_LASTHEREDATE, M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, "
		strSql = strSql & "M_HIDE_EMAIL, M_RECEIVE_EMAIL, M_LAST_IP, M_IP,M_ALLOWEMAIL) "
		strSql = strSql & "VALUES(1, '" & strAdminName & "', '" & strAdminName & "', '" & strAdminPassword & "', 'yourmail@server.com', ' ', ' ', ' ', ' ', ' ', ' ', 1, 1, 1, 3, ' ', ' ', ' ', ' ', "
		strSql = strSql & "1, '" & strCurrentDateTime & "', '" & strlhDateTime & "', '" & strCurrentDateTime & "', 'Forum Admin', 0, 0, 1, '000.000.000.000', '000.000.000.000',1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT FORUM_ID FROM " & strTablePrefix & "FORUM "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "FORUM(CAT_ID, F_STATUS, F_MAIL, F_SUBJECT, F_URL, F_DESCRIPTION, F_TOPICS, F_COUNT, F_LAST_POST, "
		strSql = strSql & "F_PASSWORD_NEW, F_PRIVATEFORUMS, F_TYPE, F_IP, F_LAST_POST_AUTHOR, F_LAST_POST_TOPIC_ID, F_LAST_POST_REPLY_ID) "
		strSql = strSql & "VALUES(1, 1, 0, 'Testing Forums', '', 'This forum gives you a chance to become more familiar with how this product responds to different features and keeps testing in one place instead of posting tests all over. Happy Posting! [:)]', "
		strSql = strSql & "1, 1, '" & strCurrentDateTime & "', '', 0, 0, '000.000.000.000', 1, 1, 0) "
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS (CAT_ID, FORUM_ID, T_STATUS, T_MAIL, T_SUBJECT, T_MESSAGE, T_AUTHOR, "
		strSql = strSql & "T_REPLIES, T_UREPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID, T_ARCHIVE_FLAG) "
		strSql = strSql & "VALUES(1, 1, 1, 0, 'Welcome to Snitz Forums 2000', 'Thank you for downloading Snitz Forums 2000. We hope you enjoy this great tool to support your organization!" & CHR(13) & CHR(10) & CHR(13) & CHR(10) &"Many thanks go out to John Penfold &lt;asp@asp-dev.com&gt; and Tim Teal &lt;tteal@tealnet.com&gt; for the original source code and to all the people of Snitz Forums 2000 at http://forum.snitz.com for continued support of this product.', "
		strSql = strSql & "1, 0, 0, 0, '" & strCurrentDateTime & "', '" & strCurrentDateTime & "', 0, '000.000.000.000', 1, 0, 1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT COUNT_ID FROM " & strTablePrefix & "TOTALS "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "TOTALS (COUNT_ID, P_COUNT, T_COUNT, U_COUNT) "
		strSql = strSql & "VALUES(1,1,1,1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT B_ID FROM " & strFilterTablePrefix & "BADWORDS "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_BADWORD, B_REPLACE) VALUES ('fuck','****')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_BADWORD, B_REPLACE) VALUES (' wank',' ****')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_BADWORD, B_REPLACE) VALUES ('shit','****')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_BADWORD, B_REPLACE) VALUES ('pussy','*****')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_BADWORD, B_REPLACE) VALUES ('cunt','****')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing
	
	strSql = "SELECT SPAM_ID FROM " & strFilterTablePrefix & "SPAM_MAIL "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@gmailinator.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@dodgeit.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@spamgourmet.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@spamhole.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@jetable.org')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@bumpymail.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@mytrashmail.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('@u-yep.com')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT GROUP_ID FROM " & strTablePrefix & "GROUP_NAMES "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) VALUES ('All Categories you have access to','All Categories you have access to')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) VALUES ('Default Categories','Default Categories')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT GROUP_ID FROM " & strTablePrefix & "GROUPS "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "GROUPS (GROUP_ID, GROUP_CATID) VALUES (2,1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing
	
	on error goto 0
end if

sub ChkDBInstall()
	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = my_conn.Errors(counter).Number
		ConnErrorDescription = my_conn.Errors(counter).Description
		if ConnErrorNumber <> 0 then 
			Err_Msg = "<tr><td class=""oops"" width=""30%"">Error: " & ConnErrorNumber & "</td>"
			Err_Msg = Err_Msg & "<td>" & ConnErrorDescription & "</td></tr>"
			Err_Msg = Err_Msg & "<tr><td class=""oops"" width=""30%"">strSql: </td>"
			Err_Msg = Err_Msg & "<td>" & strSql & "</td></tr>"
			Response.Write(Err_Msg)
			intCriticalErrors = intCriticalErrors + 1
		end if
	next
	my_conn.Errors.Clear 
end sub
%>
