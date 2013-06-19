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


function EmailField(fTestString) 
	set regEx = New RegExp
	regEx.Global = true
	regEx.IgnoreCase = true
	regEx.Pattern = "^[A-Z0-9._%-]+@(?:[A-Z0-9-]+\.)+[A-Z]{2,4}$"

	retVal = regEx.Test(fTestString)

	set regEx = nothing

	if Not retVal then
		EmailField = 0
	else
		EmailField = 1
	end if
end function 


'##############################################
'##            Ranks and Stars               ##
'##############################################

function getMember_Level(fM_TITLE, fM_LEVEL, fM_POSTS)
	dim Member_Level

	Member_Level = ""
	if Trim(fM_TITLE) <> "" then
		Member_Level = fM_TITLE
	else
		select case fM_LEVEL
			case "1"  
				if (fM_POSTS < cLng(intRankLevel1)) then Member_Level = Member_Level & strRankLevel0
				if (fM_POSTS >= cLng(intRankLevel1)) and (fM_POSTS < cLng(intRankLevel2)) then Member_Level = Member_Level & strRankLevel1
				if (fM_POSTS >= cLng(intRankLevel2)) and (fM_POSTS < cLng(intRankLevel3)) then Member_Level = Member_Level & strRankLevel2
				if (fM_POSTS >= cLng(intRankLevel3)) and (fM_POSTS < cLng(intRankLevel4)) then Member_Level = Member_Level & strRankLevel3
				if (fM_POSTS >= cLng(intRankLevel4)) and (fM_POSTS < cLng(intRankLevel5)) then Member_Level = Member_Level & strRankLevel4
				if (fM_POSTS >= cLng(intRankLevel5)) then Member_Level = Member_Level & strRankLevel5
			case "2"
				Member_Level = Member_Level & strRankMod
			case "3"
				Member_Level = Member_Level & strRankAdmin
			case else  
				Member_Level = Member_Level & "Error" 
		end select
	end if
	
	getMember_Level = Member_Level
end function


function getStar_Level(fM_LEVEL, fM_POSTS)
	dim Star_Level

	select case fM_LEVEL
		case "1"
			if (fM_POSTS < cLng(intRankLevel1)) then Star_Level = ""
			if (fM_POSTS >= cLng(intRankLevel1)) and (fM_POSTS < cLng(intRankLevel2)) then Star_Level = getCurrentIcon(getStarColor(strRankColor1),"","")
			if (fM_POSTS >= cLng(intRankLevel2)) and (fM_POSTS < cLng(intRankLevel3)) then Star_Level = getCurrentIcon(getStarColor(strRankColor2),"","") & getCurrentIcon(getStarColor(strRankColor2),"","")
			if (fM_POSTS >= cLng(intRankLevel3)) and (fM_POSTS < cLng(intRankLevel4)) then Star_Level = getCurrentIcon(getStarColor(strRankColor3),"","") & getCurrentIcon(getStarColor(strRankColor3),"","") & getCurrentIcon(getStarColor(strRankColor3),"","")
			if (fM_POSTS >= cLng(intRankLevel4)) and (fM_POSTS < cLng(intRankLevel5)) then Star_Level = getCurrentIcon(getStarColor(strRankColor4),"","") & getCurrentIcon(getStarColor(strRankColor4),"","") & getCurrentIcon(getStarColor(strRankColor4),"","") & getCurrentIcon(getStarColor(strRankColor4),"","")
			if (fM_POSTS >= cLng(intRankLevel5)) then Star_Level = getCurrentIcon(getStarColor(strRankColor5),"","") & getCurrentIcon(getStarColor(strRankColor5),"","") & getCurrentIcon(getStarColor(strRankColor5),"","") & getCurrentIcon(getStarColor(strRankColor5),"","") & getCurrentIcon(getStarColor(strRankColor5),"","")
		case "2" 
			if fM_POSTS < cLng(intRankLevel1) then Star_Level = ""
			if (fM_POSTS >= cLng(intRankLevel1)) and (fM_POSTS < cLng(intRankLevel2)) then Star_Level = getCurrentIcon(getStarColor(strRankColorMod),"","")
			if (fM_POSTS >= cLng(intRankLevel2)) and (fM_POSTS < cLng(intRankLevel3)) then Star_Level = getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","")
			if (fM_POSTS >= cLng(intRankLevel3)) and (fM_POSTS < cLng(intRankLevel4)) then Star_Level = getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","")
			if (fM_POSTS >= cLng(intRankLevel4)) and (fM_POSTS < cLng(intRankLevel5)) then Star_Level = getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","")
			if (fM_POSTS >= cLng(intRankLevel5)) then Star_Level = getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","") & getCurrentIcon(getStarColor(strRankColorMod),"","")
		case "3" 
			if (fM_POSTS < cLng(intRankLevel1)) then Star_Level = ""
			if (fM_POSTS >= cLng(intRankLevel1)) and (fM_POSTS < cLng(intRankLevel2)) then Star_Level = getCurrentIcon(getStarColor(strRankColorAdmin),"","")
			if (fM_POSTS >= cLng(intRankLevel2)) and (fM_POSTS < cLng(intRankLevel3)) then Star_Level = getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","")
			if (fM_POSTS >= cLng(intRankLevel3)) and (fM_POSTS < cLng(intRankLevel4)) then Star_Level = getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","")
			if (fM_POSTS >= cLng(intRankLevel4)) and (fM_POSTS < cLng(intRankLevel5)) then Star_Level = getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","")
			if (fM_POSTS >= cLng(intRankLevel5)) then Star_Level = getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","") & getCurrentIcon(getStarColor(strRankColorAdmin),"","")

		case else  
			Star_Level = "Error"
	end select

	getStar_Level = Star_Level
end function


function getStarColor(strStarColor)
	select case strStarColor
		case "gold"   : getStarColor = strIconStarGold
		case "silver" : getStarColor = strIconStarSilver
		case "bronze" : getStarColor = strIconStarBronze
		case "orange" : getStarColor = strIconStarOrange
		case "red"    : getStarColor = strIconStarRed
		case "purple" : getStarColor = strIconStarPurple
		case "blue"   : getStarColor = strIconStarBlue
		case "cyan"   : getStarColor = strIconStarCyan
		case "green"  : getStarColor = strIconStarGreen
	end select
end function


function getSig(fUser_Name)
	'## Forum_SQL
	strSql = "SELECT M_SIG "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(fUser_Name, "SQLString") & "'"

	set rsSig = my_Conn.Execute (strSql)

	if rsSig.EOF or rsSig.BOF then
		'## Do nothing
	else
		getSig = rsSig("M_SIG")
	end if

	rsSig.close
	set rsSig = nothing
end function


function ViewSig(fUserID)
	if fUserID = -1 then
		ViewSig = 1
		exit function
	end if

	'## Forum_SQL
	strSqlv = "SELECT M_VIEW_SIG "
	strSqlv = strSqlv & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSqlv = strSqlv & " WHERE MEMBER_ID = " & cLng(fUserID)

	set rsVSig = my_Conn.Execute (strSqlv)

	if rsVSig.EOF or rsVSig.BOF then
		ViewSig = 1
	else
		ViewSig = rsVSig("M_VIEW_SIG")
	end if

	rsVSig.close
	set rsVSig = nothing
end function


function getSigDefault(fUserID)
	if fUserID = -1 then
		getSigDefault = 1
		exit function
	end if

	if Session(strCookieURL & "intSigDefault" & MemberID) = "" or IsNull(Session(strCookieURL & "intSigDefault" & MemberID)) then
		'on error resume next
		strSqld = "SELECT M_SIG_DEFAULT "
		strSqld = strSqld & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSqld = strSqld & " WHERE MEMBER_ID = " & cLng(fUserID)

		set rsSigDefault = my_Conn.Execute (strSqld)

		if rsSigDefault.EOF or rsSigDefault.BOF then
			getSigDefault = 1
			set rsSigDefault = nothing
			exit function
		else
			tmpSigDefault = rsSigDefault("M_SIG_DEFAULT")
			Session(strCookieURL & "intSigDefault" & MemberID) = tmpSigDefault
			Session(strCookieURL & "intSigDefault" & MemberID) = tmpSigDefault
		end if

		set rsSigDefault = nothing
	end if
	if Session(strCookieURL & "intSigDefault" & MemberID) <> "" then
		getSigDefault = Session(strCookieURL & "intSigDefault" & MemberID)
	else
		getSigDefault = 1
	end if
end function


Function DisplayUsersAge(fDOB)
	dtDOB = fDOB
	dtToday = FormatDateTime(strForumTimeAdjust,2)
	DisplayUsersAge = DateDiff("yyyy", dtDOB, dtToday)
	dtTmp = DateAdd("yyyy", DisplayUsersAge, dtDOB)
	if (DateDiff("d", dtToday, dtTmp) > 0) then DisplayUsersAge = DisplayUsersAge - 1
End Function


function DOBToDate(fDOB)
	'Testing for server format
	if strComp(Month("04/05/2002"),"4") = 0 then
		DOBToDate = cdate("" & Mid(fDOB, 5,2) & "/" & Mid(fDOB, 7,2) & "/" & Mid(fDOB, 1,4) & "")
	else
		DOBToDate = cdate("" & Mid(fDOB, 7,2) & "/" & Mid(fDOB, 5,2) & "/" & Mid(fDOB, 1,4) & "")
	end if
end function
%>
