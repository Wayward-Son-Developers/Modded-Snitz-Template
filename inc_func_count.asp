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


'##############################################
'##                Do Counts                 ##
'##############################################

sub doPCount()
	'## Forum_SQL - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET P_COUNT = P_COUNT + 1"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub


sub doTCount()
	'## Forum_SQL - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET T_COUNT = T_COUNT + 1"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

'Modified function to use ID of member instead of their username.
'Function still supports updating via their username, for backward compatability.
sub doUCount(sUser)
	if VarType(sUser) = 8 then 'Update using member username
		'## Forum_SQL - Update Total Post for user
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_POSTS = M_POSTS + 1 "
		strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(sUser, "SQLString") & "'"
		
	elseif VarType(sUser) = 2 or VarType(sUser) = 3 then 'Update count using member id
		'## Forum_SQL - Update Total Post for user
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_POSTS = M_POSTS + 1 "
		strSql = strSql & " WHERE MEMBER_ID = " & sUser
		
	end if
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

'Modified function to use ID of member instead of their username.
'Function still supports updating via their username, for backward compatability.
sub doULastPost(sUser)
	if VarType(sUser) = 8 then 'Update using member user name
		'## Forum_SQL - Updates the M_LASTPOSTDATE in the FORUM_MEMBERS table
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_LASTPOSTDATE = '" & DateToStr(strForumTimeAdjust) & "' "
		strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(sUser, "SQLString") & "'"
		
	elseif VarType(sUser) = 2 or VarType(sUser) = 3 then 'Update using member id
		'## Forum_SQL - Updates the M_LASTPOSTDATE in the FORUM_MEMBERS table
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_LASTPOSTDATE = '" & DateToStr(strForumTimeAdjust) & "' "
		strSql = strSql & " WHERE MEMBER_ID = " & sUser
		
	end if
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

%>
