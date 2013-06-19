<!--#INCLUDE FILE="inc_adovbs.asp"-->
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

Session.LCID = 1033 '## Do Not Edit
Response.Buffer = true

'Anti-ForumPoster code
Dim isSpambot : isSpambot = 0
agent = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
If InStr(agent, "forum poster") > 0 Then isSpamBot = 1
If InStr(agent, "fp.icontool.com") > 0 Then isSpamBot = 1
If InStr(agent, "icontool") > 0 Then isSpamBot = 1
If isSpambot = 1 Then Response.Redirect("http://" & Request.ServerVariables("REMOTE_ADDR"))
'End Anti-ForumPoster Code

Dim strDBType, strConnString, strTablePrefix, strMemberTablePrefix, strFilterTablePrefix '## Do Not Edit
Dim counter, ConnErrorNumber, ConnErrorDesc, blnSetup '## Do Not Edit

'#################################################################################
'## SELECT YOUR DATABASE TYPE AND CONNECTION TYPE (access, sqlserver or mysql)
'#################################################################################
'strDBType = "sqlserver"
 strDBType = "access"
'strDBType = "mysql"

'## Make sure to uncomment one of the strConnString lines and edit it so that it points to where your database is!
 strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("dbase/snitz_forums_2000.mdb")         '## MS Access 2000 using virtual path
'strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/USERNAME/db/snitz_forums_2000.mdb")  '## MS Access 2000 on Brinkster
'strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\dbase\snitz_forums_2000.mdb"                   '## MS Access 2000 using physical path
'strConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("dbase/snitz_forums_2000.mdb")        '## MS Access 97 using virtual path
'strConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/USERNAME/db/snitz_forums_2000.mdb") '## MS Access 97 on Brinkster
'strConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\inetpub\dbase\snitz_forums_2000.mdb"                  '## MS Access 97 using physical path
'strConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("dbase/snitz_forums_2000.accdb")      '## MS Access 2007 using virtual path
'strConnString = "Provider=SQLOLEDB;Data Source=SERVER_NAME;database=DB_NAME;uid=UID;pwd=PWD;"                           '## MS SQL Server 6.x/7.x/2000 (OLEDB connection)
'strConnString = "driver={SQL Server};server=SERVER_NAME;uid=UID;pwd=PWD;database=DB_NAME"                               '## MS SQL Server 6.x/7.x/2000 (ODBC connection)
'strConnString = "Provider=SQLNCLI;server=SERVER_NAME;database=DB_NAME;uid=UID;pwd=PWD;"                                 '## MS SQL Server 2005 w/ SQL Native Client
'strConnString = "driver={MySQL ODBC 3.51 Driver};option=16387;server=SERVER_IP;user=UID;password=PWD;DATABASE=DB_NAME;" '## MySQL w/ MyODBC v3.51
'strConnString = "driver=MySQL;server=SERVER_IP;uid=UID;pwd=PWD;database=DB_NAME"                                        '## MySQL w/ MyODBC v2.50
'strConnString = "DSN_NAME"                                                                                              '## DSN

'#################################################################################
'## The table prefixes are used to create the table names. If you are going to 
'## run multiple forums from the same database, change the table prefixes 
'## accordingly so there is no conflict between the two forums.
'#################################################################################
strTablePrefix = "FORUM_"
strMemberTablePrefix = "FORUM_"
strFilterTablePrefix = "FORUM_"  'used for BADWORDS and NAMEFILTER tables

'#################################################################################
'## If you have deleted the default Admin account, you may need to change the
'## value below.  Otherwise, it should be left unchanged. (such as with a new
'## installation)
'#################################################################################
Const intAdminMemberID = 1

'#################################################################################
'## Will probably never need to wory about this, but if you're paranoid about
'## someone spoofing the session code used to unlock admin options, you can change
'## this value to something long and different from the 'default' value.
'#################################################################################
Const strAdminCode = "15916941253"

'#################################################################################
'## intCookieDuration is the amount of days before the forum cookie expires
'## You can set it to a higher value
'## For example for one year you can set it to 365
'## (default is 30 days)
'#################################################################################
Const intCookieDuration = 30

'#################################################################################
'## strUniqueID is used in creating the cookies and other related tasks. It 
'## uniquely identifies your forum. More often than not you won't need to change
'## this unless you are trying to run more than one forum on the same domain.
'#################################################################################
Dim strUniqueID : strUniqueID = "BasePackage"

'#################################################################################
'## strIPLookup is the address of the service you want to use for looking up IP
'## addresses. It assumes the IP address is going to be appended on the end of the
'## URL. If this isn't the case then you will need to make some changes in 
'## pop_viewip.asp, inc_profile.asp (x2), admin_accounts_pending.asp, and 
'## admin_member_search.asp (x2)
'## 
'## To change, comment out the one in use and uncomment the one you want to use
'## or add your own.
'#################################################################################
Dim strIPLookup
strIPLookup = "http://samspade.org/whois/"
'strIPLookup = "http://ws.arin.net/cgi-bin/whois.pl?queryinput="


%>
<!--#INCLUDE FILE="inc_iconfiles.asp"-->
<%
'#################################################################################
'## Do Not Edit Below This Line - It could destroy your forums and lose data
'#################################################################################

Dim mLev, strLoginStatus, MemberID, strArchiveTablePrefix
Dim strVersion, strForumTitle, strCopyright, strTitleImage, strHomeURL
Dim strForumURL, strAuthType, strSetCookieToForum, strEmail, strUniqueEmail, strFilterEMailAddresses
Dim strMailMode, strMailServer, strSender, intMaxPostsToEMail, strNoMaxPostsToEMail, strDateType, strTimeAdjust
Dim strTimeType, strMoveTopicMode, strMoveNotify, strIPLogging, strPrivateForums
Dim strShowModerators, strAllowForumCode, strIMGInPosts, strAllowHTML, strNoCookies
Dim strHotTopic, intHotTopicNum, strSecureAdmin
Dim strAIM, strICQ, strMSN, strYAHOO
Dim strFullName, strPicture, strSex, strCity, strState
Dim strAge, strAgeDOB, strMinAge, strCountry, strOccupation, strBio
Dim strHobbies, strLNews, strQuote, strMarStatus, strFavLinks
Dim strRecentTopics, strAllowHideEmail, strHomepage
Dim strReqAIM, strReqICQ, strReqMSN, strReqYAHOO
Dim strReqFullName, strReqPicture, strReqSex, strReqCity, strReqState
Dim strReqAge, strReqAgeDOB, strReqHomepage, strReqCountry, strReqOccupation, strReqBio
Dim strReqHobbies, strReqLNews, strReqQuote, strReqMarStatus, strReqFavLinks
Dim strGfxButtons, strEditedByDate, strBadWordFilter, strBadWords, strDefaultFontFace, strUseExtendedProfile, strIcons
Dim strDefaultFontSize, strHeaderFontSize, strFooterFontSize, strPageBGColor, strDefaultFontColor
Dim strLinkColor, strLinkTextDecoration, strVisitedLinkColor, strVisitedTextDecoration
Dim strActiveLinkColor, strActiveTextDecoration, strHoverFontColor, strHoverTextDecoration
Dim strHeadCellColor, strHeadFontColor, strCategoryCellColor, strCategoryFontColor
Dim strForumFirstCellColor, strForumCellColor, strAltForumCellColor, strForumFontColor
Dim strForumLinkColor, strForumLinkTextDecoration, strForumVisitedLinkColor, strForumVisitedTextDecoration
Dim strForumActiveLinkColor, strForumActiveTextDecoration, strForumHoverFontColor, strForumHoverTextDecoration
Dim strTableBorderColor, strPopUpTableColor, strPopUpBorderColor, strNewFontColor, strHiLiteFontColor, strSearchHiLiteColor
Dim strTopicWidthLeft, strTopicNoWrapLeft, strTopicWidthRight, strTopicNoWrapRight, strShowRank
Dim strRankAdmin, strRankMod, strRankColorAdmin, strRankColorMod
Dim strRankLevel0, strRankLevel1, strRankLevel2, strRankLevel3, strRankLevel4, strRankLevel5
Dim strRankColor0, strRankColor1, strRankColor2, strRankColor3, strRankColor4, strRankColor5
Dim intRankLevel0, intRankLevel1, intRankLevel2, intRankLevel3, intRankLevel4, intRankLevel5
Dim strSignatures, strDSignatures, strShowStatistics, strShowImagePoweredBy, strLogonForMail
Dim strShowPaging, strShowTopicNav, strPageSize, strPageNumberSize, strForumTimeAdjust
Dim strNTGroups, strAutoLogon, strModeration, strSubscription, strArchiveState, strUserNameFilter
Dim strFloodCheck, strFloodCheckTime, strTimeLimit, strEmailVal, strProhibitNewMembers, strRequireReg, strRestrictReg
Dim strGroupCategories, strPageBGImageUrl, strImageUrl, strJumpLastPost, strStickyTopic, strShowSendToFriend
Dim strShowPrinterFriendly, strShowTimer, strTimerPhrase, strShowFormatButtons, strShowSmiliesTable, strShowQuickReply
Dim SubCount, MySubCount
Dim strSiteHeader, strSiteLeft, strSiteRight, strSiteFooter, strSiteIntegEnabled, strSiteBorder
Dim strUGView, strUGMemView, strUGModForums

strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))
If Application(strCookieURL & "ConfigLoaded")= "" Or IsNull(Application(strCookieURL & "ConfigLoaded")) Or blnSetup="Y" Then

	on error resume next

	blnLoadConfig = TRUE

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear
	Err.Clear

	my_Conn.Open strConnString
	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDesc = my_Conn.Errors(counter).Description
		If ConnErrorNumber <> 0 Then
			If blnSetup <> "Y" Then
				my_Conn.Errors.Clear
				Err.Clear
				Response.Redirect "setup.asp?RC=1&CC=1&strDBType=" & strDBType & "&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc)
			else
				blnLoadConfig = FALSE
			end if
		end if
	next

	my_Conn.Errors.Clear
	Err.Clear 

	'## if the configvariables aren't loaded into the Application object
	'## or after the admin has changed the configuration
	'## the variables get (re)loaded

	'## Forum_SQL
	strSql = "SELECT * FROM " & strTablePrefix & "CONFIG_NEW "

	set rsConfig = my_Conn.Execute (strSql)

	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		If ConnErrorNumber <> 0 Then
			If blnSetup <> "Y" Then
				my_Conn.Errors.Clear
				Err.Clear

				strSql = "SELECT C_STRVERSION, C_STRSENDER "
				strSql = strSql & " FROM " & strTablePrefix & "CONFIG "

				set rsInfo = my_Conn.Execute (StrSql)
				strVersion = rsInfo("C_STRVERSION")
				strSender = rsInfo("C_STRSENDER")

				rsInfo.Close
				set rsInfo = nothing

				if strVersion = "" then
					strSql = "SELECT C_VALUE "
					strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
					strSql = strSql & " WHERE C_VARIABLE = 'strVersion' "
					set rsInfo = my_Conn.Execute (StrSql)
					strVersion = rsInfo("C_VALUE")
					rsInfo.Close
					set rsInfo = nothing

					strSql = "SELECT C_VALUE "
					strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
					strSql = strSql & " WHERE C_VARIABLE = 'strSender' "
					set rsInfo = my_Conn.Execute (StrSql)
					strSender = rsInfo("C_VALUE")
					rsInfo.Close
					set rsInfo = nothing
				end if

				my_Conn.Close
				set my_Conn = nothing

				Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&strDBType="& strDBType & "&EC=" & ConnErrorNumber
			else
				my_Conn.Errors.Clear
				blnLoadConfig = FALSE
			end if

		end if
	next

	my_Conn.Errors.Clear
	
	Dim appVarsLoadError
	appVarsLoadError = false
	
	if blnLoadConfig then
		Application.Lock
		do while not rsConfig.EOF
			Application(strCookieURL & Trim(UCase(rsConfig("C_VARIABLE")))) = Trim(rsConfig("C_VALUE"))
			
            ' Check for errors loading the variables...
			If Err.Number <> 0 Then appVarsLoadError = true
			
			rsConfig.MoveNext
				
           ' Check again for errors, this time moving to the next record...
			If Err.Number <> 0 Then appVarsLoadError = true
		Loop

		Application.UnLock
		rsConfig.close
		
		my_Conn.Close
		set my_Conn = nothing
		
		If Not(AppVarsLoadError) Then
			Application.Lock
			Application(strCookieURL & "ConfigLoaded")= "YES"
			Application.UnLock
		Else
			Response.Write("Server Error: The Application variables are not loaded.")
			Response.End
		End If
	end if


	my_Conn.Close
	set my_Conn = nothing

	on error goto 0
	Application.Lock
	Application(strCookieURL & "ConfigLoaded")= "YES"
	Application.UnLock
End If

' ## Read the config-info from the application variables...

strVersion = Application(strCookieURL & "STRVERSION")
strForumTitle = Application(strCookieURL & "STRFORUMTITLE")
strCopyright = Application(strCookieURL & "STRCOPYRIGHT")
strTitleImage = Application(strCookieURL & "STRTITLEIMAGE")
strHomeURL = Application(strCookieURL & "STRHOMEURL")
strForumURL = Application(strCookieURL & "STRFORUMURL")
strAuthType = Application(strCookieURL & "STRAUTHTYPE")
strSetCookieToForum = Application(strCookieURL & "STRSETCOOKIETOFORUM")
strEmail = Application(strCookieURL & "STREMAIL")
strUniqueEmail = Application(strCookieURL & "STRUNIQUEEMAIL")
strMailMode = Application(strCookieURL & "STRMAILMODE")
strMailServer = Application(strCookieURL & "STRMAILSERVER")
strSender = Application(strCookieURL & "STRSENDER")
intMaxPostsToEMail = cLng(Application(strCookieURL & "INTMAXPOSTSTOEMAIL"))
strNoMaxPostsToEMail = Application(strCookieURL & "STRNOMAXPOSTSTOEMAIL")
strFilterEMailAddresses = Application(strCookieURL & "STRFILTEREMAILADDRESSES")
strDateType = Application(strCookieURL & "STRDATETYPE")
strTimeAdjust = Application(strCookieURL & "STRTIMEADJUST")
strTimeType = Application(strCookieURL & "STRTIMETYPE")
strMoveTopicMode = Application(strCookieURL & "STRMOVETOPICMODE")
strMoveNotify = Application(strCookieURL & "STRMOVENOTIFY")
strIPLogging = Application(strCookieURL & "STRIPLOGGING")
strPrivateForums = Application(strCookieURL & "STRPRIVATEFORUMS")
strShowModerators = Application(strCookieURL & "STRSHOWMODERATORS")
strAllowForumCode = Application(strCookieURL & "STRALLOWFORUMCODE")
strIMGInPosts = Application(strCookieURL & "STRIMGINPOSTS")
strAllowHTML = Application(strCookieURL & "STRALLOWHTML")
strNoCookies = Application(strCookieURL & "STRNOCOOKIES")
strSecureAdmin = Application(strCookieURL & "STRSECUREADMIN")
strHotTopic = Application(strCookieURL & "STRHOTTOPIC")
intHotTopicNum = cLng(Application(strCookieURL & "INTHOTTOPICNUM"))
strAIM = Application(strCookieURL & "STRAIM")
strICQ = Application(strCookieURL & "STRICQ")
strMSN = Application(strCookieURL & "STRMSN")
strYAHOO = Application(strCookieURL & "STRYAHOO")
strFullName = Application(strCookieURL & "STRFULLNAME")
strPicture = Application(strCookieURL & "STRPICTURE")
strSex = Application(strCookieURL & "STRSEX")
strCity = Application(strCookieURL & "STRCITY")
strState = Application(strCookieURL & "STRSTATE")
strAge = Application(strCookieURL & "STRAGE")
strAgeDOB = Application(strCookieURL & "STRAGEDOB")
strMinAge = cInt(Application(strCookieURL & "STRMINAGE"))
strCountry = Application(strCookieURL & "STRCOUNTRY")
strOccupation = Application(strCookieURL & "STROCCUPATION")
strBio = Application(strCookieURL & "STRBIO")
strHobbies = Application(strCookieURL & "STRHOBBIES")
strLNews = Application(strCookieURL & "STRLNEWS")
strQuote = Application(strCookieURL & "STRQUOTE")
strMarStatus = Application(strCookieURL & "STRMARSTATUS")
strFavLinks = Application(strCookieURL & "STRFAVLINKS")
strRecentTopics = Application(strCookieURL & "STRRECENTTOPICS")
strAllowHideEmail = "1" '##not yet used !
strHomepage = Application(strCookieURL & "STRHOMEPAGE")
strSignatures = Application(strCookieURL & "STRSIGNATURES")
strDSignatures = Application(strCookieURL & "STRDSIGNATURES")
strUseExtendedProfile = (cLng(strSignatures) + cLng(strBio) + cLng(strHobbies) + cLng(strLNews) + cLng(strRecentTopics) + cLng(strPicture) + cLng(strQuote)) > 0
strUseExtendedProfile = strUseExtendedProfile or ((cLng(strAIM) + cLng(strICQ) + cLng(strMSN) + cLng(strYAHOO) + (cLng(strFullName)*2) + cLng(strSex) + cLng(strCity) + cLng(strState) + cLng(strAge) + cLng(strCountry) + cLng(strOccupation) + (cLng(strFavLinks)*2)) > 5)
strReqAIM = Application(strCookieURL & "STRREQAIM")
strReqICQ = Application(strCookieURL & "STRREQICQ")
strReqMSN = Application(strCookieURL & "STRREQMSN")
strReqYAHOO = Application(strCookieURL & "STRREQYAHOO")
strReqFullName = Application(strCookieURL & "STRREQFULLNAME")
strReqPicture = Application(strCookieURL & "STRREQPICTURE")
strReqSex = Application(strCookieURL & "STRREQSEX")
strReqCity = Application(strCookieURL & "STRREQCITY")
strReqState = Application(strCookieURL & "STRREQSTATE")
strReqAge = Application(strCookieURL & "STRREQAGE")
strReqAgeDOB = Application(strCookieURL & "STRREQAGEDOB")
strReqHomepage = Application(strCookieURL & "STRREQHOMEPAGE")
strReqCountry = Application(strCookieURL & "STRREQCOUNTRY")
strReqOccupation = Application(strCookieURL & "STRREQOCCUPATION")
strReqBio = Application(strCookieURL & "STRREQBIO")
strReqHobbies = Application(strCookieURL & "STRREQHOBBIES")
strReqLNews = Application(strCookieURL & "STRREQLNEWS")
strReqQuote = Application(strCookieURL & "STRREQQUOTE")
strReqMarStatus = Application(strCookieURL & "STRREQMARSTATUS")
strReqFavLinks = Application(strCookieURL & "STRREQFAVLINKS")
strIcons = Application(strCookieURL & "STRICONS")
strGfxButtons = Application(strCookieURL & "STRGFXBUTTONS")
strEditedByDate = Application(strCookieURL & "STREDITEDBYDATE")
strBadWordFilter = Application(strCookieURL & "STRBADWORDFILTER")
strBadWords = Application(strCookieURL & "STRBADWORDS")
strUserNameFilter = Application(strCookieURL & "STRUSERNAMEFILTER")
strDefaultFontFace = Application(strCookieURL & "STRDEFAULTFONTFACE")
strDefaultFontSize = Application(strCookieURL & "STRDEFAULTFONTSIZE")
strHeaderFontSize = Application(strCookieURL & "STRHEADERFONTSIZE")
strFooterFontSize = Application(strCookieURL & "STRFOOTERFONTSIZE")
strPageBGColor = Application(strCookieURL & "STRPAGEBGCOLOR")
strDefaultFontColor = Application(strCookieURL & "STRDEFAULTFONTCOLOR")
strLinkColor = Application(strCookieURL & "STRLINKCOLOR")
strLinkTextDecoration = Application(strCookieURL & "STRLINKTEXTDECORATION")
strVisitedLinkColor = Application(strCookieURL & "STRVISITEDLINKCOLOR")
strVisitedTextDecoration = Application(strCookieURL & "STRVISITEDTEXTDECORATION")
strActiveLinkColor = Application(strCookieURL & "STRACTIVELINKCOLOR")
strActiveTextDecoration = Application(strCookieURL & "STRACTIVETEXTDECORATION")
strHoverFontColor = Application(strCookieURL & "STRHOVERFONTCOLOR")
strHoverTextDecoration = Application(strCookieURL & "STRHOVERTEXTDECORATION")
strHeadCellColor = Application(strCookieURL & "STRHEADCELLCOLOR")
strHeadFontColor = Application(strCookieURL & "STRHEADFONTCOLOR")
strCategoryCellColor = Application(strCookieURL & "STRCATEGORYCELLCOLOR")
strCategoryFontColor = Application(strCookieURL & "STRCATEGORYFONTCOLOR")
strForumFirstCellColor = Application(strCookieURL & "STRFORUMFIRSTCELLCOLOR")
strForumCellColor = Application(strCookieURL & "STRFORUMCELLCOLOR")
strAltForumCellColor = Application(strCookieURL & "STRALTFORUMCELLCOLOR")
strForumFontColor = Application(strCookieURL & "STRFORUMFONTCOLOR")
strForumLinkColor = Application(strCookieURL & "STRFORUMLINKCOLOR")
strForumLinkTextDecoration = Application(strCookieURL & "STRFORUMLINKTEXTDECORATION")
strForumVisitedLinkColor = Application(strCookieURL & "STRFORUMVISITEDLINKCOLOR")
strForumVisitedTextDecoration = Application(strCookieURL & "STRFORUMVISITEDTEXTDECORATION")
strForumActiveLinkColor = Application(strCookieURL & "STRFORUMACTIVELINKCOLOR")
strForumActiveTextDecoration = Application(strCookieURL & "STRFORUMACTIVETEXTDECORATION")
strForumHoverFontColor = Application(strCookieURL & "STRFORUMHOVERFONTCOLOR")
strForumHoverTextDecoration = Application(strCookieURL & "STRFORUMHOVERTEXTDECORATION")
strTableBorderColor = Application(strCookieURL & "STRTABLEBORDERCOLOR")
strPopUpTableColor = Application(strCookieURL & "STRPOPUPTABLECOLOR")
strPopUpBorderColor = Application(strCookieURL & "STRPOPUPBORDERCOLOR")
strNewFontColor = Application(strCookieURL & "STRNEWFONTCOLOR")
strHiLiteFontColor = Application(strCookieURL & "STRHILITEFONTCOLOR")
strSearchHiLiteColor = Application(strCookieURL & "STRSEARCHHILITECOLOR")
strTopicWidthLeft = Application(strCookieURL & "STRTOPICWIDTHLEFT")
strTopicNoWrapLeft = Application(strCookieURL & "STRTOPICNOWRAPLEFT")
strTopicWidthRight = Application(strCookieURL & "STRTOPICWIDTHRIGHT")
strTopicNoWrapRight = Application(strCookieURL & "STRTOPICNOWRAPRIGHT")
strShowRank = Application(strCookieURL & "STRSHOWRANK")
strRankAdmin = Application(strCookieURL & "STRRANKADMIN")
strRankMod = Application(strCookieURL & "STRRANKMOD")
strRankLevel0 = Application(strCookieURL & "STRRANKLEVEL0")
strRankLevel1 = Application(strCookieURL & "STRRANKLEVEL1")
strRankLevel2 = Application(strCookieURL & "STRRANKLEVEL2")
strRankLevel3 = Application(strCookieURL & "STRRANKLEVEL3")
strRankLevel4 = Application(strCookieURL & "STRRANKLEVEL4")
strRankLevel5 = Application(strCookieURL & "STRRANKLEVEL5")
strRankColorAdmin = Application(strCookieURL & "STRRANKCOLORADMIN")
strRankColorMod = Application(strCookieURL & "STRRANKCOLORMOD")
strRankColor0 = Application(strCookieURL & "STRRANKCOLOR0")
strRankColor1 = Application(strCookieURL & "STRRANKCOLOR1")
strRankColor2 = Application(strCookieURL & "STRRANKCOLOR2")
strRankColor3 = Application(strCookieURL & "STRRANKCOLOR3")
strRankColor4 = Application(strCookieURL & "STRRANKCOLOR4")
strRankColor5 = Application(strCookieURL & "STRRANKCOLOR5")
intRankLevel0 = Application(strCookieURL & "INTRANKLEVEL0")
intRankLevel1 = Application(strCookieURL & "INTRANKLEVEL1")
intRankLevel2 = Application(strCookieURL & "INTRANKLEVEL2")
intRankLevel3 = Application(strCookieURL & "INTRANKLEVEL3")
intRankLevel4 = Application(strCookieURL & "INTRANKLEVEL4")
intRankLevel5 = Application(strCookieURL & "INTRANKLEVEL5")
strShowStatistics = Application(strCookieURL & "STRSHOWSTATISTICS")
strShowImagePoweredBy = Application(strCookieURL & "STRSHOWIMAGEPOWEREDBY")
strLogonForMail = Application(strCookieURL & "STRLOGONFORMAIL")
strShowPaging = Application(strCookieURL & "STRSHOWPAGING")
strShowTopicNav = Application(strCookieURL & "STRSHOWTOPICNAV")
strPageSize = Application(strCookieURL & "STRPAGESIZE")
strPageNumberSize = Application(strCookieURL & "STRPAGENUMBERSIZE")
strForumTimeAdjust = DateAdd("h", strTimeAdjust , Now())
strNTGroups = Application(strCookieURL & "STRNTGROUPS")
strAutoLogon = Application(strCookieURL & "STRAUTOLOGON")
strModeration = Application(strCookieURL & "STRMODERATION")
strSubscription = Application(strCookieURL & "STRSUBSCRIPTION")
strArchiveState = Application(strCookieURL & "STRARCHIVESTATE")
strFloodCheck = Application(strCookieURL & "STRFLOODCHECK")
strFloodCheckTime = Application(strCookieURL & "STRFLOODCHECKTIME")
strEmailVal = Application(strCookieURL & "STREMAILVAL")
strPageBGImageUrl = Application(strCookieURL & "STRPAGEBGIMAGEURL")
strImageUrl = Application(strCookieURL & "STRIMAGEURL")
strJumpLastPost = Application(strCookieURL & "STRJUMPLASTPOST")
strStickyTopic = Application(strCookieURL & "STRSTICKYTOPIC")
strShowSendToFriend = Application(strCookieURL & "STRSHOWSENDTOFRIEND")
strShowPrinterFriendly = Application(strCookieURL & "STRSHOWPRINTERFRIENDLY")
strProhibitNewMembers = Application(strCookieURL & "STRPROHIBITNEWMEMBERS")
strRequireReg = Application(strCookieURL & "STRREQUIREREG")
strRestrictReg = Application(strCookieURL & "STRRESTRICTREG")
strGroupCategories = Application(strCookieURL & "STRGROUPCATEGORIES")
strShowTimer = Application(strCookieURL & "STRSHOWTIMER")
strTimerPhrase = Application(strCookieURL & "STRTIMERPHRASE")
strShowFormatButtons = Application(strCookieURL & "STRSHOWFORMATBUTTONS")
strShowSmiliesTable = Application(strCookieURL & "STRSHOWSMILIESTABLE")
strShowQuickReply = Application(strCookieURL & "STRSHOWQUICKREPLY")
strSiteHeader = Application(strCookieURL & "STRSITEHEADER")
strSiteLeft = Application(strCookieURL & "STRSITELEFT")
strSiteRight = Application(strCookieURL & "STRSITERIGHT")
strSiteFooter = Application(strCookieURL & "STRSITEFOOTER")
strSiteIntegEnabled = Application(strCookieURL & "STRSITEINTEGENABLED")
strSiteBorder = Application(strCookieURL & "STRSITEBORDER")
strUGView = Application(strCookieURL & "STRUGVIEW")
strUGMemView = Application(strCookieURL & "STRUGMEMVIEW")
strUGModForums = Application(strCookieURL & "STRUGMODFORUMS")

if strSecureAdmin = "0" then
	Session(strCookieURL & "Approval") = strAdminCode
end if

if strAuthType = "db" then
	strDBNTSQLName = "M_NAME"
	strAutoLogon = "0"
	strNTGroups = "0"
else
	strDBNTSQLName = "M_USERNAME"
end if
%>
