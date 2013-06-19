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

if (strDBType = "sqlserver") then
	if strSQL_Server = "SQL6" then
		strN = ""
	else
		strN = "n"
	end if

	my_Conn.Errors.Clear
	on error resume next

	strSql = "CREATE TABLE " & strTablePrefix & "CATEGORY ( "
	strSql = strSql & "CAT_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "CAT_STATUS smallint NULL , "
	strSql = strSql & "CAT_NAME " & strN & "varchar (100) NULL , "
	strSql = strSql & "CAT_MODERATION int NULL CONSTRAINT " & strTablePrefix & "SnitzC1020 DEFAULT 0, "
	strSql = strSql & "CAT_SUBSCRIPTION int NULL CONSTRAINT " & strTablePrefix & "SnitzC1021 DEFAULT 0, "
	strSql = strSql & "CAT_ORDER int NULL CONSTRAINT " & strTablePrefix & "SnitzC1022 DEFAULT 1 )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "CONFIG_NEW ( "
	strSql = strSql & "ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "C_VARIABLE " & strN & "varchar (255) NULL , "
	strSql = strSql & "C_VALUE " & strN & "varchar (255) NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "FORUM ( "
	strSql = strSql & "CAT_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "F_STATUS smallint NULL , "
	strSql = strSql & "F_MAIL smallint NULL , "
	strSql = strSql & "F_SUBJECT " & strN & "varchar (100) NULL , "
	strSql = strSql & "F_URL " & strN & "varchar (255) NULL , "
	strSql = strSql & "F_DESCRIPTION " & strN & "text NULL , "
	strSql = strSql & "F_TOPICS int NULL , "
	strSql = strSql & "F_COUNT int NULL , "
	strSql = strSql & "F_LAST_POST " & strN & "varchar (14) NULL , "
	strSql = strSql & "F_PASSWORD_NEW " & strN & "varchar (255) NULL , "
	strSql = strSql & "F_PRIVATEFORUMS int NULL , "
	strSql = strSql & "F_TYPE smallint NULL , "
	strSql = strSql & "F_IP " & strN & "varchar (50) NULL,  "
	strSql = strSql & "F_LAST_POST_AUTHOR int NULL, "
	strSql = strSql & "F_LAST_POST_TOPIC_ID int NULL, "
	strSql = strSql & "F_LAST_POST_REPLY_ID int NULL, "
	strSql = strSql & "F_A_TOPICS int NULL , "
	strSql = strSql & "F_A_COUNT int NULL , "
	strSql = strSql & "F_DEFAULTDAYS int NULL DEFAULT 30 , "
	strSql = strSql & "F_COUNT_M_POSTS smallint NULL DEFAULT 1 , "
	strSql = strSql & "F_MODERATION int NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1018 DEFAULT 0, "
	strSql = strSql & "F_SUBSCRIPTION int NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1019 DEFAULT 0, "
	strSql = strSql & "F_ORDER int NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1025 DEFAULT 1, "
	strSql = strSql & "F_L_ARCHIVE " & strN & "varchar (14) NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1026 DEFAULT '', "
	strSql = strSql & "F_ARCHIVE_SCHED int NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1027 DEFAULT 30, "
	strSql = strSql & "F_L_DELETE " & strN & "varchar (14) NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1028 DEFAULT '', "
	strSql = strSql & "F_DELETE_SCHED int NOT NULL CONSTRAINT " & strTablePrefix & "SnitzC1029 DEFAULT 365 ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS ( "
	strSql = strSql & "MEMBER_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "M_STATUS smallint NULL , "
	strSql = strSql & "M_NAME " & strN & "varchar (75) NULL DEFAULT '' , "
	strSql = strSql & "M_USERNAME " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_PASSWORD " & strN & "varchar (65) NULL DEFAULT '' , "
	strSql = strSql & "M_EMAIL " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_COUNTRY " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_HOMEPAGE " & strN & "varchar (255) NULL DEFAULT '' , "
	strSql = strSql & "M_SIG " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_VIEW_SIG smallint NULL DEFAULT 1 , "
	strSql = strSql & "M_SIG_DEFAULT smallint NULL DEFAULT 1 , "
	strSql = strSql & "M_DEFAULT_VIEW int NULL , "
	strSql = strSql & "M_LEVEL smallint NULL , "
	strSql = strSql & "M_AIM " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_ICQ " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_MSN " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_YAHOO " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_POSTS int NULL DEFAULT '0' , "
	strSql = strSql & "M_DATE " & strN & "varchar (14) NULL , "
	strSql = strSql & "M_LASTHEREDATE " & strN & "varchar (14) NULL DEFAULT '' , "
	strSql = strSql & "M_LASTPOSTDATE " & strN & "varchar (14) NULL DEFAULT '' , "
	strSql = strSql & "M_TITLE " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_SUBSCRIPTION smallint NULL , "
	strSql = strSql & "M_HIDE_EMAIL smallint NULL , "
	strSql = strSql & "M_RECEIVE_EMAIL smallint NULL , "
	strSql = strSql & "M_LAST_IP " & strN & "varchar (50) NULL , "
	strSql = strSql & "M_IP " & strN & "varchar (50) NULL , "
	strSql = strSql & "M_FIRSTNAME " & strN & "varchar (100) NULL CONSTRAINT " & strTablePrefix & "SnitzC0369 DEFAULT '' ,"
	strSql = strSql & "M_LASTNAME " & strN & "varchar (100) NULL CONSTRAINT " & strTablePrefix & "SnitzC0370 DEFAULT '' ,"
	strSql = strSql & "M_OCCUPATION " & strN & "varchar (255) NULL CONSTRAINT " & strTablePrefix & "SnitzC0371 DEFAULT '' ,"
	strSql = strSql & "M_SEX " & strN & "varchar (50) NULL CONSTRAINT " & strTablePrefix & "SnitzC0372 DEFAULT '' , "
	strSql = strSql & "M_AGE " & strN & "varchar (10) NULL CONSTRAINT " & strTablePrefix & "SnitzC0373 DEFAULT '' , "
	strSql = strSql & "M_DOB " & strN & "varchar (8) NULL DEFAULT '' , "
	strSql = strSql & "M_HOBBIES " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_LNEWS " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_QUOTE " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_BIO " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_MARSTATUS " & strN & "varchar (100) NULL CONSTRAINT " & strTablePrefix & "SnitzC0374 DEFAULT '' ,"
	strSql = strSql & "M_LINK1 " & strN & "varchar (255) NULL CONSTRAINT " & strTablePrefix & "SnitzC0375 DEFAULT '' ,"
	strSql = strSql & "M_LINK2 " & strN & "varchar (255) NULL CONSTRAINT " & strTablePrefix & "SnitzC0376 DEFAULT '' , "
	strSql = strSql & "M_CITY " & strN & "varchar (100) NULL CONSTRAINT " & strTablePrefix & "SnitzC0377 DEFAULT '' , "
	strSql = strSql & "M_STATE " & strN & "varchar (100) NULL CONSTRAINT " & strTablePrefix & "SnitzC0379 DEFAULT '' , "
	strSql = strSql & "M_PHOTO_URL " & strN & "varchar (255) NULL CONSTRAINT " & strTablePrefix & "SnitzC0378 DEFAULT '' , "
	strSql = strSql & "M_KEY " & strN & "varchar (32) NULL DEFAULT '' , "
	strSql = strSql & "M_NEWEMAIL " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_PWKEY " & strN & "varchar (32) NULL DEFAULT '' , "
	strSql = strSql & "M_SHA256 smallint NULL DEFAULT 1 , "
	strSql = strSql & "M_ALLOWEMAIL smallint NULL DEFAULT 0 )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ( "
	strSql = strSql & "MEMBER_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "M_STATUS smallint NULL , "
	strSql = strSql & "M_NAME " & strN & "varchar (75) NULL DEFAULT '' , "
	strSql = strSql & "M_USERNAME " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_PASSWORD " & strN & "varchar (65) NULL DEFAULT '' , "
	strSql = strSql & "M_EMAIL " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_COUNTRY " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_HOMEPAGE " & strN & "varchar (255) NULL DEFAULT '' , "
	strSql = strSql & "M_SIG " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_VIEW_SIG smallint NULL DEFAULT 1 , "
	strSql = strSql & "M_SIG_DEFAULT smallint NULL DEFAULT 1 , "
	strSql = strSql & "M_DEFAULT_VIEW int NULL , "
	strSql = strSql & "M_LEVEL smallint NULL , "
	strSql = strSql & "M_AIM " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_ICQ " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_MSN " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_YAHOO " & strN & "varchar (150) NULL DEFAULT '' , "
	strSql = strSql & "M_POSTS int NULL DEFAULT '0' , "
	strSql = strSql & "M_DATE " & strN & "varchar (14) NULL , "
	strSql = strSql & "M_LASTHEREDATE " & strN & "varchar (14) NULL DEFAULT '' , "
	strSql = strSql & "M_LASTPOSTDATE " & strN & "varchar (14) NULL DEFAULT '' , "
	strSql = strSql & "M_TITLE " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_SUBSCRIPTION smallint NULL , "
	strSql = strSql & "M_HIDE_EMAIL smallint NULL DEFAULT 0, "
	strSql = strSql & "M_RECEIVE_EMAIL smallint NULL DEFAULT 1, "
	strSql = strSql & "M_LAST_IP " & strN & "varchar (50) NULL , "
	strSql = strSql & "M_IP " & strN & "varchar (50) NULL , "
	strSql = strSql & "M_FIRSTNAME " & strN & "varchar (100) NULL DEFAULT '' ,"
	strSql = strSql & "M_LASTNAME " & strN & "varchar (100) NULL DEFAULT '' ,"
	strSql = strSql & "M_OCCUPATION " & strN & "varchar (255) NULL DEFAULT '' ,"
	strSql = strSql & "M_SEX " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_AGE " & strN & "varchar (10) NULL DEFAULT '' , "
	strSql = strSql & "M_DOB " & strN & "varchar (8) NULL DEFAULT '' , "
	strSql = strSql & "M_HOBBIES " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_LNEWS " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_QUOTE " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_BIO " & strN & "text NULL DEFAULT '' , "
	strSql = strSql & "M_MARSTATUS " & strN & "varchar (100) NULL DEFAULT '' ,"
	strSql = strSql & "M_LINK1 " & strN & "varchar (255) NULL DEFAULT '' ,"
	strSql = strSql & "M_LINK2 " & strN & "varchar (255) NULL DEFAULT '' , "
	strSql = strSql & "M_CITY " & strN & "varchar (100) NULL DEFAULT '' , "
	strSql = strSql & "M_STATE " & strN & "varchar (100) NULL DEFAULT '' , "
	strSql = strSql & "M_PHOTO_URL " & strN & "varchar (255) NULL DEFAULT '' , "
	strSql = strSql & "M_KEY " & strN & "varchar (32) NULL DEFAULT '' , "
	strSql = strSql & "M_NEWEMAIL " & strN & "varchar (50) NULL DEFAULT '' , "
	strSql = strSql & "M_PWKEY " & strN & "varchar (32) NULL DEFAULT '' , "
	strSql = strSql & "M_APPROVE smallint NULL DEFAULT '' , "
	strSql = strSql & "M_SHA256 smallint NULL DEFAULT 1 , "
	strSql = strSql & "M_ALLOWEMAIL smallint NULL DEFAULT 0 )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "MODERATOR ( "
	strSql = strSql & "MOD_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "FORUM_ID int NULL , "
	strSql = strSql & "MEMBER_ID int NULL , "
	strSql = strSql & "MOD_TYPE smallint NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "REPLY ( "
	strSql = strSql & "CAT_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int NOT NULL , "
	strSql = strSql & "TOPIC_ID int NOT NULL , "
	strSql = strSql & "REPLY_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "R_MAIL smallint NULL , "
	strSql = strSql & "R_AUTHOR int NULL , "
	strSql = strSql & "R_MESSAGE " & strN & "text NULL , "
	strSql = strSql & "R_DATE " & strN & "varchar (14) NULL , "
	strSql = strSql & "R_IP " & strN & "varchar (50) NULL, "
	strSql = strSql & "R_STATUS smallint NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT 0 , "
	strSql = strSql & "R_LAST_EDIT " & strN & "varchar (14) NULL , " 
	strSql = strSql & "R_LAST_EDITBY int NULL , "
	strSql = strSql & "R_SIG smallint NULL ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "TOPICS ( "
	strSql = strSql & "CAT_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int NOT NULL , "
	strSql = strSql & "TOPIC_ID int IDENTITY (1, 1) NOT NULL , "
	strSql = strSql & "T_STATUS smallint NULL , "
	strSql = strSql & "T_MAIL smallint NULL , "
	strSql = strSql & "T_SUBJECT " & strN & "varchar (100) NULL , "
	strSql = strSql & "T_MESSAGE " & strN & "text NULL , "
	strSql = strSql & "T_AUTHOR int NULL , "
	strSql = strSql & "T_REPLIES int NULL , "
	strSql = strSql & "T_UREPLIES int NULL , "
	strSql = strSql & "T_VIEW_COUNT int NULL , "
	strSql = strSql & "T_LAST_POST " & strN & "varchar (14) NULL , "
	strSql = strSql & "T_DATE " & strN & "varchar (14) NULL , "
	strSql = strSql & "T_LAST_POSTER int NULL , "
	strSql = strSql & "T_IP " & strN & "varchar (50) NULL , " 
	strSql = strSql & "T_LAST_POST_AUTHOR int NULL , "
	strSql = strSql & "T_LAST_POST_REPLY_ID int NULL , "
	strSql = strSql & "T_ARCHIVE_FLAG int NULL , " 
	strSql = strSql & "T_LAST_EDIT " & strN & "varchar (14) NULL , " 
	strSql = strSql & "T_LAST_EDITBY int NULL , " 
	strSql = strSql & "T_STICKY smallint NULL DEFAULT 0, " 
	strSql = strSql & "T_SIG smallint NULL DEFAULT 0) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "TOTALS ( "
	strSql = strSql & "COUNT_ID smallint NOT NULL , "
	strSql = strSql & "P_COUNT int NULL , "
	strSql = strSql & "P_A_COUNT int NULL , "
	strSql = strSql & "T_COUNT int NULL , "
	strSql = strSql & "T_A_COUNT int NULL , "
	strSql = strSql & "U_COUNT int NULL ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
	strSql = strSql & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY NONCLUSTERED (MEMBER_ID, FORUM_ID) ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "SUBSCRIPTIONS ("
	strSql = strSql & "SUBSCRIPTION_ID INT IDENTITY NOT NULL, MEMBER_ID INT NOT NULL, "
	strSql = strSql & "CAT_ID INT NOT NULL, TOPIC_ID INT NOT NULL, FORUM_ID INT NOT NULL) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
	strSql = strSql & "CAT_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int NOT NULL , "
	strSql = strSql & "TOPIC_ID int NOT NULL , "
	strSql = strSql & "T_STATUS smallint NULL , "
	strSql = strSql & "T_MAIL smallint NULL , "
	strSql = strSql & "T_SUBJECT " & strN & "varchar (100) NULL , "
	strSql = strSql & "T_MESSAGE " & strN & "text NULL , "
	strSql = strSql & "T_AUTHOR int NULL , "
	strSql = strSql & "T_REPLIES int NULL , "
	strSql = strSql & "T_UREPLIES int NULL , "
	strSql = strSql & "T_VIEW_COUNT int NULL , "
	strSql = strSql & "T_LAST_POST " & strN & "varchar (14) NULL , "
	strSql = strSql & "T_DATE " & strN & "varchar (14) NULL , "
	strSql = strSql & "T_LAST_POSTER int NULL , "
	strSql = strSql & "T_IP " & strN & "varchar (50) NULL , " 
	strSql = strSql & "T_LAST_POST_AUTHOR int NULL , "
	strSql = strSql & "T_LAST_POST_REPLY_ID int NULL , "
	strSql = strSql & "T_ARCHIVE_FLAG int NULL , " 
	strSql = strSql & "T_LAST_EDIT " & strN & "varchar (14) NULL , " 
	strSql = strSql & "T_LAST_EDITBY int NULL , " 
	strSql = strSql & "T_STICKY smallint NULL , " 
	strSql = strSql & "T_SIG smallint NULL ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
	strSql = strSql & "CAT_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int NOT NULL , "
	strSql = strSql & "TOPIC_ID int NOT NULL , "
	strSql = strSql & "REPLY_ID int NOT NULL , "
 	strSql = strSql & "R_STATUS smallint NULL , "
	strSql = strSql & "R_MAIL smallint NULL , "
	strSql = strSql & "R_AUTHOR int NULL , "
	strSql = strSql & "R_MESSAGE " & strN & "text NULL , "
	strSql = strSql & "R_DATE " & strN & "varchar (14) NULL , "
	strSql = strSql & "R_IP " & strN & "varchar (50) NULL , "
	strSql = strSql & "R_LAST_EDIT " & strN & "varchar (14) NULL , " 
	strSql = strSql & "R_LAST_EDITBY int NULL , "
	strSql = strSql & "R_SIG smallint NULL ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strFilterTablePrefix & "BADWORDS ( "
	strSql = strSql & "B_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "B_BADWORD " & strN & "varchar (50) NULL , "
	strSql = strSql & "B_REPLACE " & strN & "varchar (50) NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
	strSql = strSql & "SPAM_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "SPAM_SERVER " & strN & "varchar (255) NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "CUSTOM_POLICY ( "
	strSql = strSql & "CP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "CP_MODE smallint NOT NULL , "
	strSql = strSql & "CP_CONTENT " & strN & "text NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "FAQ ( "
	strSql = strSql & "F_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "F_FAQ_QUESTION " & strN & "varchar (255) NULL , "
	strSql = strSql & "F_FAQ_ANSWER " & strN & "text NULL , "
	strSql = strSql & "F_FAQ_ORDER int NULL DEFAULT 1 , "
	strSql = strSql & "F_FAQ_CATEGORY int NULL DEFAULT 1 , "
	strSql = strSql & "F_FAQ_TYPE int NULL DEFAULT 0)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "FAQ_CATEGORY ( "
	strSql = strSql & "FCAT_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "FCAT_TITLE " & strN & "varchar (100) NULL , "
	strSql = strSql & "FCAT_ORDER int NULL DEFAULT 1 , "
	strSql = strSql & "FCAT_LEVEL int NULL)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "USERGROUPS ( "
	strSql = strSql & "USERGROUP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "USERGROUP_NAME " & strN & "varchar(100) NOT NULL , "
	strSql = strSql & "USERGROUP_DESC " & strN & "text NULL , "
	strSql = strSql & "USERGROUP_TYPE int NOT NULL DEFAULT 0 , "
	strSql = strSql & "AUTOJOIN int NOT NULL DEFAULT 0 , "
	strSql = strSql & "MEM_HIDE int NOT NULL DEFAULT 0 , "
	strSql = strSql & "MOD_HIDE int NOT NULL DEFAULT 0)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "USERGROUP_USERS ( "
	strSql = strSql & "USERGROUP_ID int NOT NULL , "
	strSql = strSql & "MEMBER_ID int NOT NULL , "
	strSql = strSql & "MEMBER_TYPE int NOT NULL DEFAULT 1)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "ALLOWED_USERGROUPS ( "
	strSql = strSql & "USERGROUP_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int NOT NULL , "
	strSql = strSql & "PERMS int NOT NULL DEFAULT 0)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strFilterTablePrefix & "NAMEFILTER ( "
	strSql = strSql & "N_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "N_NAME " & strN & "varchar (75) NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "GROUP_NAMES ( "
	strSql = strSql & "GROUP_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "GROUP_NAME " & strN & "varchar (50) NULL , "
	strSql = strSql & "GROUP_DESCRIPTION " & strN & "varchar (255) NULL , "
	strSql = strSql & "GROUP_ICON " & strN & "varchar (255) NULL , "
	strSql = strSql & "GROUP_IMAGE " & strN & "varchar (255) NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "GROUPS ( "
	strSql = strSql & "GROUP_KEY int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
	strSql = strSql & "GROUP_ID int NULL , "
	strSql = strSql & "GROUP_CATID int NULL )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strTablePrefix & "CATEGORY WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC1 DEFAULT 1 FOR CAT_STATUS, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC2 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & " (CAT_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()


	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG_NEW WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC38 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & " (ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strTablePrefix & "FORUM WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC39 DEFAULT 0 FOR CAT_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC40 DEFAULT 1 FOR F_STATUS, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC41 DEFAULT 0 FOR F_MAIL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC42 DEFAULT 0 FOR F_TOPICS, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC43 DEFAULT 0 FOR F_COUNT, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC44 DEFAULT '' FOR F_LAST_POST, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC45 DEFAULT 0 FOR F_PRIVATEFORUMS, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC46 DEFAULT 0 FOR F_TYPE, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC47 DEFAULT '000.000.000.000' FOR F_IP, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC48 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & "(CAT_ID,	FORUM_ID )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC49 DEFAULT 1 FOR M_STATUS, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC50 DEFAULT 1 FOR M_DEFAULT_VIEW, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC51 DEFAULT 1 FOR M_LEVEL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC52 DEFAULT '' FOR M_DATE, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC53 DEFAULT 0 FOR M_SUBSCRIPTION, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC54 DEFAULT 0 FOR M_HIDE_EMAIL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC55 DEFAULT 1 FOR M_RECEIVE_EMAIL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC56 DEFAULT '000.000.000.000' FOR M_LAST_IP, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC57 DEFAULT '000.000.000.000' FOR M_IP, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC58 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & "(MEMBER_ID)"

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strTablePrefix & "MODERATOR WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC59 DEFAULT 0 FOR FORUM_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC60 DEFAULT 0 FOR MEMBER_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC61 DEFAULT 0 FOR MOD_TYPE, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC62 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & " (MOD_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strTablePrefix & "REPLY WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC63 DEFAULT 0 FOR CAT_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC64 DEFAULT 0 FOR FORUM_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC65 DEFAULT 0 FOR TOPIC_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC66 DEFAULT 0 FOR R_MAIL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC67 DEFAULT 0 FOR R_AUTHOR, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC68 DEFAULT '' FOR R_DATE, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC69 DEFAULT '000.000.000.000' FOR R_IP, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC70 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & "(CAT_ID,	FORUM_ID, TOPIC_ID,	REPLY_ID )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strTablePrefix & "TOPICS WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC71 DEFAULT 0 FOR CAT_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC72 DEFAULT 0 FOR FORUM_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC73 DEFAULT 1 FOR T_STATUS, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC74 DEFAULT 0 FOR T_MAIL, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC75 DEFAULT 0 FOR T_AUTHOR, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC76 DEFAULT 0 FOR T_REPLIES, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC77 DEFAULT 0 FOR T_VIEW_COUNT, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC78 DEFAULT '' FOR T_LAST_POST, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC79 DEFAULT '' FOR T_DATE, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC80 DEFAULT 0 FOR T_LAST_POSTER, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC81 DEFAULT '000.000.000.000' FOR T_IP, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC82 DEFAULT 0 FOR T_ARCHIVE_FLAG, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC83 PRIMARY KEY  NONCLUSTERED "
	strSql = strSql & "(CAT_ID, FORUM_ID, TOPIC_ID )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "ALTER TABLE " & strTablePrefix & "TOTALS WITH NOCHECK ADD "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC84 DEFAULT 0 FOR COUNT_ID, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC85 DEFAULT 0 FOR P_COUNT, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC86 DEFAULT 0 FOR T_COUNT, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC87 DEFAULT 0 FOR U_COUNT, "
	strSql = strSql & "CONSTRAINT " & strTablePrefix & "SnitzC88 PRIMARY KEY  NONCLUSTERED " 
	strSql = strSql & "(COUNT_ID) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "CATEGORY_CAT_ID ON " & strTablePrefix & "CATEGORY(CAT_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "CATEGORY_CAT_STATUS ON " & strTablePrefix & "CATEGORY(CAT_STATUS)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "CAT_ID ON " & strTablePrefix & "FORUM(CAT_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "F_CAT ON " & strTablePrefix & "FORUM(CAT_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "FORUM_ID ON " & strTablePrefix & "FORUM(FORUM_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strMemberTablePrefix & "MEMBERS_MEMBER_ID ON " &strMemberTablePrefix & "MEMBERS(MEMBER_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "MODERATOR_FORUM_ID ON " & strTablePrefix & "MODERATOR(FORUM_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "MODERATORS_MEMBER_ID ON " & strTablePrefix & "MODERATOR(MEMBER_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "REPLY_R_AUTHOR ON " & strTablePrefix & "REPLY(R_AUTHOR)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "REPLY_CATFORTOP_ID ON " & strTablePrefix & "REPLY(CAT_ID, FORUM_ID, TOPIC_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "FORUM_ID ON " & strTablePrefix & "REPLY(FORUM_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "REPLY_ID ON " & strTablePrefix & "REPLY(REPLY_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "REPLY_TOPIC_ID ON " & strTablePrefix & "REPLY(TOPIC_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "TOPICS_CAT_ID_FORUM_ID ON " & strTablePrefix & "TOPICS(CAT_ID, FORUM_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "TOPICS_T_AUTHOR ON " & strTablePrefix & "TOPICS(T_AUTHOR)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "TOPICS_CAT_ID ON " & strTablePrefix & "TOPICS(CAT_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "FORUM_ID ON " & strTablePrefix & "TOPICS(FORUM_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "TOPICS_TOPIC_ID ON " & strTablePrefix & "TOPICS(TOPIC_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE INDEX " & strTablePrefix & "TOPICS_CAT_FOR_TOP ON " & strTablePrefix & "TOPICS(CAT_ID, FORUM_ID, TOPIC_ID)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

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
		strSql = "INSERT " & strTablePrefix & "CATEGORY(CAT_STATUS, CAT_NAME) VALUES(1, 'Snitz Forums 2000')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS WHERE M_LEVEL = 3"
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT " & strMemberTablePrefix & "MEMBERS (M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, "
		strSql = strSql & "M_HOMEPAGE, M_LINK1, M_LINK2, M_PHOTO_URL, M_SIG, M_VIEW_SIG, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_ICQ, M_MSN, M_YAHOO, "
		strSql = strSql & "M_POSTS, M_DATE, M_LASTHEREDATE, M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, "
		strSql = strSql & "M_HIDE_EMAIL, M_RECEIVE_EMAIL, M_LAST_IP, M_IP) "
		strSql = strSql & " VALUES(1, '" & strAdminName & "', '" & strAdminName & "', '" & strAdminPassword & "', 'yourmail@server.com', ' ', ' ', ' ', ' ', ' ', ' ', 1, 1, 3, ' ', ' ', ' ', ' ', "
		strSql = strSql & " 1, '" & strCurrentDateTime & "', '" & strlhDateTime & "', '" & strCurrentDateTime & "', 'Forum Admin', 0, 0, 1, '000.000.000.000', '000.000.000.000')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT FORUM_ID FROM " & strTablePrefix & "FORUM "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT " & strTablePrefix & "FORUM(CAT_ID, F_STATUS, F_MAIL, F_SUBJECT, F_URL, F_DESCRIPTION, F_TOPICS, F_COUNT, F_LAST_POST, "
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
		strSql = "INSERT " & strTablePrefix & "TOPICS (CAT_ID, FORUM_ID, T_STATUS, T_MAIL, T_SUBJECT, T_MESSAGE, T_AUTHOR, "
		strSql = strSql & "T_REPLIES, T_UREPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID, T_ARCHIVE_FLAG) "
		strSql = strSql & "VALUES(1, 1, 1, 0, 'Welcome to Snitz Forums 2000', 'Thank you for downloading the Snitz Forums 2000. We hope you enjoy this great tool to support your organization!" & CHR(13) & CHR(10) & CHR(13) & CHR(10) &"Many thanks go out to John Penfold &lt;asp@asp-dev.com&gt; and Tim Teal &lt;tteal@tealnet.com&gt; for the original source code and to all the people of Snitz Forums 2000 at http://forum.snitz.com for continued support of this product.', "
		strSql = strSql & "1, 0, 0, 0, '" & strCurrentDateTime & "', '" & strCurrentDateTime & "', 0, '000.000.000.000', 1, 0, 1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT COUNT_ID FROM " & strTablePrefix & "TOTALS "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT " & strTablePrefix & "TOTALS (COUNT_ID, P_COUNT, T_COUNT, U_COUNT) "
		strSql = strSql & "VALUES(1,1,1,1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT CP_ID FROM " & strFilterTablePrefix & "CUSTOM_POLICY "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "CUSTOM_POLICY "
		strSql = strSql & " (CP_MODE, CP_CONTENT) VALUES (0,'Your custom policy here')"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT FCAT_ID FROM " & strFilterTablePrefix & "FAQ_CATEGORY "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "FAQ_CATEGORY "
		strSql = strSql & "(FCAT_TITLE,FCAT_ORDER,FCAT_LEVEL) VALUES ('General',1,0)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
	end if
	rs.close
	set rs = nothing

	strSql = "SELECT F_ID FROM " & strFilterTablePrefix & "FAQ "
	set rs = my_Conn.Execute(strSql)
	if rs.EOF then
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Do I have to register?','',1,1,1)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('How can I use smilies and images?','',1,1,2)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I add a hyperlink to my messages?','',1,1,3)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I change the format of my text?','',1,1,4)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What are Moderators?','',1,1,5)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Are cookies used?','',1,1,6)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What are active topics?','',1,1,7)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I edit my own posts?','',1,1,8)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I attach files?','',1,1,9)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I search?','',1,1,10)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I edit my profile?','',1,1,11)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I attach my own signature to my posts?','',1,1,12)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Are there any censor features?','',1,1,13)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What do I do if I forget my Password?','',1,1,14)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Can I be notified by e-mail when there are new posts?','',1,1,15)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What does it mean if a forum has Moderation enabled?','',1,1,16)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What is COPPA?','',1,1,17)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('Where can I get my own copy of this Forum?','',1,1,18)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('What is RSS?','',1,1,19)"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		ChkDBInstall()
		strSql = "INSERT INTO " & strTablePrefix & "FAQ "
		strSql = strSql & "(F_FAQ_QUESTION,F_FAQ_ANSWER,F_FAQ_ORDER,F_FAQ_CATEGORY,F_FAQ_TYPE) VALUES ('How do I access an RSS Feed?','',1,1,20)"
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
		strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_BADWORD, B_REPLACE) VALUES ('wank','****')"
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
