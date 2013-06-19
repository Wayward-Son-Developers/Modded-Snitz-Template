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

if (strDBType = "mysql") then
	on error resume next

	if (my_Conn.State = 0) Then
		my_Conn.Open strConnString
	end If

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "CATEGORY "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "CATEGORY ( "
	strSql = strSql & "CAT_ID INT (11) NOT NULL auto_increment, "
	strSql = strSql & "CAT_STATUS SMALLINT (6) DEFAULT '1' NOT NULL , "
	strSql = strSql & "CAT_NAME VARCHAR (100) DEFAULT '', "
	strSql = strSql & "CAT_MODERATION int (11) DEFAULT '0', "
	strSql = strSql & "CAT_SUBSCRIPTION int (11) DEFAULT '0', "
	strSql = strSql & "CAT_ORDER int (11) DEFAULT '1', "
	strSql = strSql & "PRIMARY KEY (CAT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "CATEGORY_CAT_ID(CAT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "CATEGORY_CAT_STATUS (CAT_STATUS) )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "CONFIG "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "CONFIG_NEW "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "CONFIG_NEW ( "
	strSql = strSql & "ID int (11) NOT NULL auto_increment, "
	strSql = strSql & "C_VARIABLE VARCHAR (255) , "
	strSql = strSql & "C_VALUE VARCHAR (255),  "
	strSql = strSql & "PRIMARY KEY (ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "FORUM "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "FORUM ( "
	strSql = strSql & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
	strSql = strSql & "FORUM_ID smallint (6) NOT NULL auto_increment, "
	strSql = strSql & "F_STATUS smallint (6) DEFAULT '1', "
	strSql = strSql & "F_MAIL smallint (6) DEFAULT '1' , "
	strSql = strSql & "F_SUBJECT VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "F_URL VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "F_DESCRIPTION TEXT , "
	strSql = strSql & "F_TOPICS DOUBLE DEFAULT '0' , "
	strSql = strSql & "F_COUNT DOUBLE DEFAULT '0' , "
	strSql = strSql & "F_LAST_POST VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "F_PASSWORD_NEW VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "F_PRIVATEFORUMS int (11) DEFAULT '0' , "
	strSql = strSql & "F_TYPE smallint (6) DEFAULT '0' , "
	strSql = strSql & "F_IP VARCHAR (50) DEFAULT '000.000.000.000' ,  "
	strSql = strSql & "F_LAST_POST_AUTHOR int (11) DEFAULT '1' ,  "
	strSql = strSql & "F_LAST_POST_TOPIC_ID int (11) DEFAULT '0' ,  "
	strSql = strSql & "F_LAST_POST_REPLY_ID int (11) DEFAULT '0' ,  "
	strSql = strSql & "F_MODERATION int (11) DEFAULT '0', "
	strSql = strSql & "F_SUBSCRIPTION int (11) DEFAULT '0' , "
	strSql = strSql & "F_ORDER int (11) DEFAULT '1' , "
	strSql = strSql & "F_DEFAULTDAYS int (11) DEFAULT '30' , "
	strSql = strSql & "F_COUNT_M_POSTS smallint (6) DEFAULT '1' , "
	strSql = strSql & "F_L_ARCHIVE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "F_ARCHIVE_SCHED int (11) DEFAULT '30' , "
	strSql = strSql & "F_L_DELETE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "F_DELETE_SCHED int (11) DEFAULT '365' , "
	strSql = strSql & "F_A_TOPICS DOUBLE DEFAULT '0' , "
	strSql = strSql & "F_A_COUNT DOUBLE DEFAULT '0' , "
	strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "FORUM_FORUM_ID(FORUM_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "FORUM_CAT_ID(CAT_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strMemberTablePrefix & "MEMBERS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS ( "
	strSql = strSql & "MEMBER_ID int (11) NOT NULL AUTO_INCREMENT, "
	strSql = strSql & "M_STATUS smallint (6) DEFAULT '0' , "
	strSql = strSql & "M_NAME VARCHAR (75) DEFAULT '' , "
	strSql = strSql & "M_USERNAME VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_PASSWORD VARCHAR (65) DEFAULT '' , "
	strSql = strSql & "M_EMAIL VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_COUNTRY VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_HOMEPAGE VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_SIG TEXT , "
	strSql = strSql & "M_VIEW_SIG smallint (6) NULL DEFAULT '1' , "
	strSql = strSql & "M_SIG_DEFAULT smallint (6) NULL DEFAULT '1' , "
	strSql = strSql & "M_DEFAULT_VIEW int (11) DEFAULT '1' , "
	strSql = strSql & "M_LEVEL smallint (6) DEFAULT '1' , "
	strSql = strSql & "M_AIM VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_ICQ VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_MSN VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_YAHOO VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_POSTS int (11) DEFAULT '0' , "
	strSql = strSql & "M_DATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "M_LASTHEREDATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "M_LASTPOSTDATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "M_TITLE VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_SUBSCRIPTION smallint (6) DEFAULT '0' , "
	strSql = strSql & "M_HIDE_EMAIL smallint (6) DEFAULT '0' , "
	strSql = strSql & "M_RECEIVE_EMAIL smallint (6) DEFAULT '1' , "
	strSql = strSql & "M_LAST_IP VARCHAR (50) DEFAULT '000.000.000.000' , "
	strSql = strSql & "M_IP VARCHAR (50) DEFAULT '000.000.000.000' , "
	strSql = strSql & "M_FIRSTNAME VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_LASTNAME VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_OCCUPATION VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_SEX VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_AGE VARCHAR (10) DEFAULT '' , "
	strSql = strSql & "M_DOB VARCHAR (8) DEFAULT '' , "
	strSql = strSql & "M_HOBBIES TEXT , "
	strSql = strSql & "M_LNEWS TEXT , "
	strSql = strSql & "M_QUOTE TEXT , "
	strSql = strSql & "M_BIO TEXT , "
	strSql = strSql & "M_MARSTATUS VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_LINK1 VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_LINK2 VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_CITY VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_STATE VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_PHOTO_URL VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_KEY VARCHAR (32) DEFAULT '' , "
	strSql = strSql & "M_NEWEMAIL VARCHAR (50) DEFAULT '', "
	strSql = strSql & "M_PWKEY VARCHAR (32) DEFAULT '' , "
	strSql = strSql & "M_SHA256 smallint (6) DEFAULT '1' ,"
	strSql = strSql & "M_ALLOWEMAIL smallint (6) NULL DEFAULT '0' , "
	strSql = strSql & "PRIMARY KEY (MEMBER_ID), "
	strSql = strSql & "KEY " & strMemberTablePrefix & "MEMBERS_MEMBER_ID (MEMBER_ID) ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strMemberTablePrefix & "MEMBERS_PENDING "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ( "
	strSql = strSql & "MEMBER_ID int (11) NOT NULL AUTO_INCREMENT, "
	strSql = strSql & "M_STATUS smallint (6) DEFAULT '0' , "
	strSql = strSql & "M_NAME VARCHAR (75) DEFAULT '' , "
	strSql = strSql & "M_USERNAME VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_PASSWORD VARCHAR (65) DEFAULT '' , "
	strSql = strSql & "M_EMAIL VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_COUNTRY VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_HOMEPAGE VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_SIG TEXT , "
	strSql = strSql & "M_VIEW_SIG smallint (6) NULL DEFAULT '1' , "
	strSql = strSql & "M_SIG_DEFAULT smallint (6) NULL DEFAULT '1' , "
	strSql = strSql & "M_DEFAULT_VIEW int (11) DEFAULT '1' , "
	strSql = strSql & "M_LEVEL smallint (6) DEFAULT '1' , "
	strSql = strSql & "M_AIM VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_ICQ VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_MSN VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_YAHOO VARCHAR (150) DEFAULT '' , "
	strSql = strSql & "M_POSTS int (11) DEFAULT '0' , "
	strSql = strSql & "M_DATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "M_LASTHEREDATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "M_LASTPOSTDATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "M_TITLE VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_SUBSCRIPTION smallint (6) DEFAULT '0' , "
	strSql = strSql & "M_HIDE_EMAIL smallint (6) DEFAULT '0' , "
	strSql = strSql & "M_RECEIVE_EMAIL smallint (6) DEFAULT '1' , "
	strSql = strSql & "M_LAST_IP VARCHAR (50) DEFAULT '000.000.000.000' , "
	strSql = strSql & "M_IP VARCHAR (50) DEFAULT '000.000.000.000' , "
	strSql = strSql & "M_FIRSTNAME VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_LASTNAME VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_OCCUPATION VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_SEX VARCHAR (50) DEFAULT '' , "
	strSql = strSql & "M_AGE VARCHAR (10) DEFAULT '' , "
	strSql = strSql & "M_DOB VARCHAR (8) DEFAULT '' , "
	strSql = strSql & "M_HOBBIES TEXT , "
	strSql = strSql & "M_LNEWS TEXT , "
	strSql = strSql & "M_QUOTE TEXT , "
	strSql = strSql & "M_BIO TEXT , "
	strSql = strSql & "M_MARSTATUS VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_LINK1 VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_LINK2 VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_CITY VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_STATE VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "M_PHOTO_URL VARCHAR (255) DEFAULT '' , "
	strSql = strSql & "M_KEY VARCHAR (32) DEFAULT '' , "
	strSql = strSql & "M_NEWEMAIL VARCHAR (50) DEFAULT '', "
	strSql = strSql & "M_PWKEY VARCHAR (32) DEFAULT '' , "
	strSql = strSql & "M_APPROVE smallint (6) , "
	strSql = strSql & "M_SHA256 smallint (6) DEFAULT '1' , "
	strSql = strSql & "M_ALLOWEMAIL smallint (6) NULL DEFAULT '0' , "
	strSql = strSql & "PRIMARY KEY (MEMBER_ID) ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "MODERATOR "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "MODERATOR ( "
	strSql = strSql & "MOD_ID int (11) NOT NULL auto_increment, "
	strSql = strSql & "FORUM_ID int (11) DEFAULT '1' , "
	strSql = strSql & "MEMBER_ID int (11) DEFAULT '1'  , "
	strSql = strSql & "MOD_TYPE smallint (6) DEFAULT '0', "
	strSql = strSql & "PRIMARY KEY (MOD_ID))"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "REPLY "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "REPLY ( "
	strSql = strSql & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
	strSql = strSql & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
	strSql = strSql & "TOPIC_ID int (11) DEFAULT '1' NOT NULL , "
	strSql = strSql & "REPLY_ID int (11) NOT NULL auto_increment, "
	strSql = strSql & "R_MAIL smallint (6) DEFAULT '0' , "
	strSql = strSql & "R_AUTHOR int (11) DEFAULT '1' , "
	strSql = strSql & "R_MESSAGE text , "
	strSql = strSql & "R_DATE VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "R_IP VARCHAR (50) DEFAULT '000.000.000.000', "
	strSql = strSql & "R_STATUS smallint (6) DEFAULT '0', "
	strSql = strSql & "R_LAST_EDIT VARCHAR (14) , " 
	strSql = strSql & "R_LAST_EDITBY int (11) , "
	strSql = strSql & "R_SIG smallint (6) DEFAULT '0', "
	strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "REPLY_CATFORTOPREPL(CAT_ID,FORUM_ID,TOPIC_ID, REPLY_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "REPLY_REP_ID(REPLY_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "REPLY_CAT_ID(CAT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "REPLY_FORUM_ID(FORUM_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "REPLY_TOPIC_ID (TOPIC_ID) )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "TOPICS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "TOPICS ( "
	strSql = strSql & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
	strSql = strSql & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
	strSql = strSql & "TOPIC_ID int (11) NOT NULL auto_increment, "
	strSql = strSql & "T_STATUS smallint (6) DEFAULT '1' , "
	strSql = strSql & "T_MAIL smallint (6) DEFAULT '0' , "
	strSql = strSql & "T_SUBJECT VARCHAR (100) DEFAULT '' , "
	strSql = strSql & "T_MESSAGE text , "
	strSql = strSql & "T_AUTHOR int (11) DEFAULT '1' , "
	strSql = strSql & "T_REPLIES DOUBLE DEFAULT '0' , "
	strSql = strSql & "T_UREPLIES int (11) DEFAULT '0' , "
	strSql = strSql & "T_VIEW_COUNT int (11) DEFAULT '0' , "
	strSql = strSql & "T_LAST_POST VARCHAR (14) DEFAULT '' , "
	strSql = strSql & "T_DATE VARCHAR (14) DEFAULT '', "
	strSql = strSql & "T_LAST_POSTER int (11) DEFAULT '1', "
	strSql = strSql & "T_IP VARCHAR (50) DEFAULT '000.000.000.000', " 
	strSql = strSql & "T_LAST_POST_AUTHOR int (11) DEFAULT '1', "
	strSql = strSql & "T_LAST_POST_REPLY_ID int (11) DEFAULT '0', "
	strSql = strSql & "T_ARCHIVE_FLAG int (11) DEFAULT '1', "
	strSql = strSql & "T_LAST_EDIT VARCHAR (14) , " 
	strSql = strSql & "T_LAST_EDITBY int (11) , " 
	strSql = strSql & "T_STICKY smallint (6) DEFAULT '0', "
	strSql = strSql & "T_SIG smallint (6) DEFAULT '0', "
	strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "TOPIC_CATFORTOP(CAT_ID,FORUM_ID,TOPIC_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "TOPIC_CAT_ID(CAT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "TOPIC_FORUM_ID(FORUM_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "TOPIC_TOPIC_ID (TOPIC_ID) )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "TOTALS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "TOTALS ( "
	strSql = strSql & "COUNT_ID smallint (6) NOT NULL auto_increment, "
	strSql = strSql & "P_COUNT int (11) DEFAULT '0' , "
	strSql = strSql & "T_COUNT int (11) DEFAULT '0'  , "
	strSql = strSql & "P_A_COUNT int (11) DEFAULT '0' , "
	strSql = strSql & "T_A_COUNT int (11) DEFAULT '0' , "
	strSql = strSql & "U_COUNT int (11) DEFAULT '0' , "
	strSql = strSql & "PRIMARY KEY (COUNT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "TOTALS_COUNT_ID (COUNT_ID) ) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "ALLOWED_MEMBERS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
	strSql = strSql & "MEMBER_ID INT (11) NOT NULL, FORUM_ID smallint (6) NOT NULL , "
	strSql = strSql & "PRIMARY KEY (MEMBER_ID, FORUM_ID) )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "SUBSCRIPTIONS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "SUBSCRIPTIONS ("
	strSql = strSql & "SUBSCRIPTION_ID INT (11) NOT NULL auto_increment, MEMBER_ID INT NOT NULL, "
	strSql = strSql & "CAT_ID INT NOT NULL, TOPIC_ID INT NOT NULL, FORUM_ID INT NOT NULL, "
	strSql = strSql & "KEY " & strTablePrefix & "SUBSCRIPTIONS_SUB_ID(SUBSCRIPTION_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "A_TOPICS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "A_TOPICS ( "
	strSql = strSql & "CAT_ID int (11) NOT NULL , "
	strSql = strSql & "FORUM_ID int (11) NOT NULL , "
	strSql = strSql & "TOPIC_ID int (11) NOT NULL, "
	strSql = strSql & "T_STATUS smallint (6) , "
	strSql = strSql & "T_MAIL smallint (6) , "
	strSql = strSql & "T_SUBJECT VARCHAR (100) , "
	strSql = strSql & "T_MESSAGE text , "
	strSql = strSql & "T_AUTHOR int (11) , "
	strSql = strSql & "T_REPLIES DOUBLE , "
	strSql = strSql & "T_UREPLIES int (11) , "
	strSql = strSql & "T_VIEW_COUNT int (11) , "
	strSql = strSql & "T_LAST_POST VARCHAR (14) , "
	strSql = strSql & "T_DATE VARCHAR (14) , "
	strSql = strSql & "T_LAST_POSTER int (11) , "
	strSql = strSql & "T_IP VARCHAR (50) , " 
	strSql = strSql & "T_LAST_POST_AUTHOR int (11) , "
	strSql = strSql & "T_LAST_POST_REPLY_ID int (11) , "
	strSql = strSql & "T_ARCHIVE_FLAG int (11) , "
	strSql = strSql & "T_LAST_EDIT VARCHAR (14) , " 
	strSql = strSql & "T_LAST_EDITBY int (11) , " 
	strSql = strSql & "T_STICKY smallint (6) , "
	strSql = strSql & "T_SIG smallint (6) , "
	strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_TOPIC_CATFORTOP(CAT_ID,FORUM_ID,TOPIC_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_TOPIC_CAT_ID(CAT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_TOPIC_FORUM_ID(FORUM_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_TOPIC_TOPIC_ID (TOPIC_ID) )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "A_REPLY "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "A_REPLY ( "
	strSql = strSql & "CAT_ID int (11) NOT NULL , "
	strSql = strSql & "FORUM_ID int (11) NOT NULL , "
	strSql = strSql & "TOPIC_ID int (11) NOT NULL , "
	strSql = strSql & "REPLY_ID int (11) NOT NULL, "
	strSql = strSql & "R_STATUS smallint (6) , "
	strSql = strSql & "R_MAIL smallint (6) , "
	strSql = strSql & "R_AUTHOR int (11) , "
	strSql = strSql & "R_MESSAGE text , "
	strSql = strSql & "R_DATE VARCHAR (14) , "
	strSql = strSql & "R_IP VARCHAR (50) , "
	strSql = strSql & "R_LAST_EDIT VARCHAR (14) , " 
	strSql = strSql & "R_LAST_EDITBY int (11) , " 
	strSql = strSql & "R_SIG smallint (6) , "
	strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_REPLY_CATFORTOPREPL(CAT_ID,FORUM_ID,TOPIC_ID, REPLY_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_REPLY_REP_ID(REPLY_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_REPLY_CAT_ID(CAT_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_REPLY_FORUM_ID(FORUM_ID), "
	strSql = strSql & "KEY " & strTablePrefix & "A_REPLY_TOPIC_ID (TOPIC_ID) )"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strFilterTablePrefix & "BADWORDS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strFilterTablePrefix & "BADWORDS ( "
	strSql = strSql & "B_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "B_BADWORD VARCHAR (50), "
	strSql = strSql & "B_REPLACE VARCHAR (50),  "
	strSql = strSql & "PRIMARY KEY (B_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "DROP TABLE IF EXISTS " & strFilterTablePrefix & "SPAM_MAIL "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
	strSql = strSql & "SPAM_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "SPAM_SERVER VARCHAR (255) NULL , "
	strSql = strSql & "PRIMARY KEY (SPAM_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "DROP TABLE IF EXISTS " & strFilterTablePrefix & "CUSTOM_POLICY "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "CUSTOM_POLICY ( "
	strSql = strSql & "CP_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "CP_MODE smallint (6) NOT NULL , "
	strSql = strSql & "CP_CONTENT text NULL , "
	strSql = strSql & "PRIMARY KEY (CP_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "DROP TABLE IF EXISTS " & strFilterTablePrefix & "FAQ "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "FAQ ( "
	strSql = strSql & "F_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "F_FAQ_QUESTION varchar (255) NULL , "
	strSql = strSql & "F_FAQ_ANSWER text NULL , "
	strSql = strSql & "F_FAQ_ORDER int NULL DEFAULT 1 , "
	strSql = strSql & "F_FAQ_CATEGORY int NULL DEFAULT 1 , "
	strSql = strSql & "F_FAQ_TYPE int NULL DEFAULT 0 , "
	strSql = strSql & "PRIMARY KEY (F_ID))"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "DROP TABLE IF EXISTS " & strFilterTablePrefix & "FAQ_CATEGORY "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "CREATE TABLE " & strTablePrefix & "FAQ_CATEGORY ( "
	strSql = strSql & "FCAT_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "FCAT_TITLE varchar (100) NULL , "
	strSql = strSql & "FCAT_ORDER int NULL DEFAULT 1 , "
	strSql = strSql & "FCAT_LEVEL int NULL DEFAULT 1 , "
	strSql = strSql & "PRIMARY KEY (FCAT_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "DROP TABLE IF EXISTS " & strFilterTablePrefix & "NAMEFILTER "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strFilterTablePrefix & "NAMEFILTER ( "
	strSql = strSql & "N_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "N_NAME VARCHAR (75),  "
	strSql = strSql & "PRIMARY KEY (N_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "USERGROUPS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "USERGROUPS ( "
	strSql = strSql & "USERGROUP_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "USERGROUP_NAME varchar(100) NOT NULL , "
	strSql = strSql & "USERGROUP_DESC text NULL , "
	strSql = strSql & "USERGROUP_TYPE int NOT NULL DEFAULT 0 , "
	strSql = strSql & "AUTOJOIN int NOT NULL DEFAULT 0 , "
	strSql = strSql & "MEM_HIDE int NOT NULL DEFAULT 0 , "
	strSql = strSql & "MOD_HIDE int NOT NULL DEFAULT 0 , "
	strSql = strSql & "PRIMARY KEY (USERGROUP_ID))"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "USERGROUP_USERS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "USERGROUP_USERS ( "
	strSql = strSql & "USERGROUP_ID int NOT NULL , "
	strSql = strSql & "MEMBER_ID int NOT NULL , "
	strSql = strSql & "MEMBER_TYPE int NOT NULL DEFAULT 1)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "ALLOWED_USERGROUPS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "ALLOWED_USERGROUPS ( "
	strSql = strSql & "USERGROUP_ID int NOT NULL , "
	strSql = strSql & "FORUM_ID int NOT NULL , "
	strSql = strSql & "PERMS int NOT NULL DEFAULT 0)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "GROUP_NAMES "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "GROUP_NAMES ( "
	strSql = strSql & "GROUP_ID int (11) NOT NULL auto_increment , "
	strSql = strSql & "GROUP_NAME VARCHAR (50) NULL , "
	strSql = strSql & "GROUP_DESCRIPTION VARCHAR (255) NULL , "
	strSql = strSql & "GROUP_ICON VARCHAR (255) NULL , "
	strSql = strSql & "GROUP_IMAGE VARCHAR (255) NULL , "
	strSql = strSql & "PRIMARY KEY (GROUP_ID)) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "GROUPS "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "CREATE TABLE " & strTablePrefix & "GROUPS ( "
	strSql = strSql & "GROUP_KEY int (11) NOT NULL auto_increment , "
	strSql = strSql & "GROUP_ID int (11) NULL , "
	strSql = strSql & "GROUP_CATID int (11) NULL , "
	strSql = strSql & "PRIMARY KEY (GROUP_KEY)) "
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

	strSql = "INSERT INTO " & strTablePrefix & "CATEGORY(CAT_STATUS, CAT_NAME) VALUES(1, 'Snitz Forums 2000')"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERS (M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, "
	strSql = strSql & "M_HOMEPAGE, M_LINK1, M_LINK2, M_PHOTO_URL, M_SIG, M_VIEW_SIG, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_ICQ, M_MSN, M_YAHOO, "
	strSql = strSql & "M_POSTS, M_DATE, M_LASTHEREDATE, M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, "
	strSql = strSql & "M_HIDE_EMAIL, M_RECEIVE_EMAIL, M_LAST_IP, M_IP) "
	strSql = strSql & "VALUES(1, '" & strAdminName & "', '" & strAdminName & "', '" & strAdminPassword & "', 'yourmail@server.com', ' ', ' ', ' ', ' ', ' ', ' ', 1, 1, 3, ' ', ' ', ' ', ' ', "
	strSql = strSql & "1, '" & strCurrentDateTime & "', '" & strlhDateTime & "', '" & strCurrentDateTime & "', 'Forum Admin', '0', '0', 1, '000.000.000.000', '000.000.000.000')"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "INSERT INTO " & strTablePrefix & "FORUM(CAT_ID, F_STATUS, F_MAIL, F_SUBJECT, F_URL, F_DESCRIPTION, F_TOPICS, F_COUNT, F_LAST_POST, "
	strSql = strSql & "F_PASSWORD_NEW, F_PRIVATEFORUMS, F_TYPE, F_IP, F_LAST_POST_AUTHOR, F_LAST_POST_TOPIC_ID, F_LAST_POST_REPLY_ID) "
	strSql = strSql & "VALUES(1, 1, '0', 'Testing Forums', '', 'This forum gives you a chance to become more familiar with how this product responds to different features and keeps testing in one place instead of posting tests all over. Happy Posting! ', "
	strSql = strSql & "1, 1, '" & strCurrentDateTime & "', '', '0', '0', '000.000.000.000', 1, 1, 0) "
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "INSERT INTO " & strTablePrefix & "TOPICS (CAT_ID, FORUM_ID, T_STATUS, T_MAIL, T_SUBJECT, T_MESSAGE, T_AUTHOR, "
	strSql = strSql & "T_REPLIES, T_UREPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID, T_ARCHIVE_FLAG) "
	strSql = strSql & "VALUES(1, 1, 1, '0', 'Welcome to Snitz Forums 2000', 'Thank you for downloading the Snitz Forums 2000. We hope you enjoy this great tool to support your organization!" & CHR(13) & CHR(10) & CHR(13) & CHR(10) &"Many thanks go out to John Penfold <asp@asp-dev.com> and Tim Teal <tteal@tealnet.com> for the original source code and to all the people of Snitz Forums 2000 at http://forum.snitz.com for continued support of this product.', "
	strSql = strSql & "1, '0', '0', '0', '" & strCurrentDateTime & "', '" & strCurrentDateTime & "', '0', '000.000.000.000', 1, 0, 1)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "INSERT INTO " & strTablePrefix & "TOTALS (COUNT_ID, P_COUNT, T_COUNT, U_COUNT) "
	strSql = strSql & "VALUES(1,1,1,1)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "INSERT INTO " & strTablePrefix & "CUSTOM_POLICY "
	strSql = strSql & " (CP_MODE, CP_CONTENT) VALUES (0,'Your custom policy here')"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
	strSql = "INSERT INTO " & strTablePrefix & "FAQ_CATEGORY "
	strSql = strSql & "(FCAT_TITLE,FCAT_ORDER,FCAT_LEVEL) VALUES ('General',1,0)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
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

	strSql = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) VALUES ('All Categories you have access to','All Categories you have access to')"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "INSERT INTO " & strTablePrefix & "GROUP_NAMES (GROUP_NAME,GROUP_DESCRIPTION) VALUES ('Default Categories','Default Categories')"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()

	strSql = "INSERT INTO " & strTablePrefix & "GROUPS (GROUP_ID, GROUP_CATID) VALUES (2,1)"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	ChkDBInstall()
	
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
	
	on error goto 0
end if

sub ChkDBInstall()
	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = my_Conn.Errors(counter).Number
		ConnErrorDescription = my_Conn.Errors(counter).Description

		if ConnErrorNumber <> 0 or Err.Number <> 0 then 
			Err_Msg = "<tr><td class=""oops"" width=""30%"">Error: " & ConnErrorNumber & "</td>"
			Err_Msg = Err_Msg & "<td>" & ConnErrorDescription & "</td></tr>"
			Err_Msg = Err_Msg & "<tr><td class=""oops"" width=""30%"">strSql: </td>"
			Err_Msg = Err_Msg & "<td>" & strSql & "</td></tr>"	
			Response.Write(Err_Msg)
			intCriticalErrors = intCriticalErrors + 1
		end if
	next
	my_Conn.Errors.Clear 
end sub
%>
