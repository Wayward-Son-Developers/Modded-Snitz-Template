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


Function SetConfigValue(bUpdate, fVariable, fValue)

	' bUpdate = 1 : if it exists then overwrite with new values
	' bUpdate = 0 : if it exists then leave unchanged

	Dim strSql

	strSql = "SELECT C_VARIABLE FROM " & strTablePrefix & "CONFIG_NEW " &_
		 " WHERE C_VARIABLE = '" & fVariable & "' "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if (rs.EOF or rs.BOF) then '## New config-value
		SetConfigValue = "added"
		my_conn.execute ("INSERT INTO " & strTablePrefix & "CONFIG_NEW (C_VALUE,C_VARIABLE) VALUES ('" & fValue & "' , '" & fVariable & "')"),,adCmdText + adExecuteNoRecords
	else
		if bUpdate <> 0 then 
			SetConfigValue = "updated"
			my_conn.execute ("UPDATE " & strTablePrefix & "CONFIG_NEW SET C_VALUE = '" & fValue & "' WHERE C_VARIABLE = '" & fVariable &"'"),,adCmdText + adExecuteNoRecords
		else ' not changed
			SetConfigValue = "unchanged"
		end if
	end if

	rs.close
	set rs = nothing
end function
%>
