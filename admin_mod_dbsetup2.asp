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
<%
if Session(strCookieURL & "Approval") <> strAdminCode then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp"-->
<%
Dim strTableName
Dim fieldArray (100)
Dim idFieldName
Dim tableExists
Dim fieldExists
Dim ErrorCount

tableExists   = -2147217900
tableNotExist = -2147217865 
fieldExists   = -2147217887
ErrorCount = 0

Response.Write	"<table class=""misc"">" & vbNewLine & _
				"<tr>" & vbNewLine & _
				"<td class=""secondnav"">" & vbNewLine & _
				getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">" & strForumTitle & "</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
				getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;Alt.&nbsp;MOD&nbsp;Setup</td>" & vbNewLine & _
				"</tr>" & vbNewLine & _
				"</table>" & vbNewLine

if MemberID <> intAdminMemberID then
	Call FailMessage("<li>Only the Forum Admin can access this page</li>",True)
	WriteFooter
	Response.End
end if

on error resume next
Select case strDBType
	case "access"
		strUserDBType = "Microsoft Access"
	case "sqlserver"
		strUserDBType = "Microsoft SQL Server"
	case "mysql"
		strUserDBType = "MySQL Server"
	case else	
		Call FailMessage("<li>Your strDBType is not set. Edit your config.asp to reflect your database type.</li>",True)
		WriteFooter
		Response.End
end Select

strRqMethod = Trim(LCase(Request.Form("method")))
if strRqMethod = "process" then
	if Request.Form("Message") = "" then
		Call FailMessage("<li>You did not enter any code to process.</li>",True)
		WriteFooter
		Response.End
	end if
	
	codetoprocess = split(Request.Form("Message"), chr(13) + chr(10))
	keycnt = ubound(codetoprocess)
	
	x = 0
	strModTitle = codetoprocess(x)
	Select case uCase(strModTitle)
		case "[CREATE]","[ALTER]","[DELETE]","[INSERT]","[UPDATE]","[DROP]"
			strModTitle = "Database Update"
		case else
	end select
	Response.Write	"<table class=""admin"" width=""75%"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td>Alt. MOD Setup</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr class=""section"">" & vbNewLine & _
					"<td>" & strModTitle & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td>" & vbNewLine
	sqlVer = Request.Form("sqltype")

	do while x < keycnt
		sectionName = codetoprocess(x)
		Select case uCase(sectionName)
			case "[CREATE]","[ALTER]","[DELETE]","[INSERT]","[UPDATE]","[DROP]"
				x = x + 1
				Select case uCase(sectionName)
					case "[CREATE]" 
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						idFieldName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						CreateTables(rec)
					case "[ALTER]" 
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						AlterTables(rec)
					case "[DELETE]" 
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						DeleteValues(rec)
					case "[INSERT]" 
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						InsertValues(rec)
					case "[UPDATE]" 
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						UpdateValues(rec)
					case "[DROP]" 
						strTableName = codetoprocess(x)
						x = x + 1
						tempField = codetoprocess(x)
						DropTable()
				end select
				x = x + 1
			case else
				x = x + 1
		end select
	loop
	
	Response.write	"<p>Database setup finished. If you have questions, please post your question in the " & _
					"<a href=""http://forum.snitz.com/forum/forum.asp?FORUM_ID=94"">MOD Implementation Forum</a>.</p>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine

else

	Response.write 	"<form action=""" & Request.ServerVariables("PATH_INFO") & """ method=""post"" name=""form1"">" & vbNewLine & _
					"<input name=""method"" type=""hidden"" value=""Process"">" & vbNewLine & _
					"<table class=""admin"" width=""75%"">" & vbNewLine & _
					"<tr class=""header"">" & vbNewLine & _
					"<td colspan=""2"">Alt. MOD Setup</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Database Type:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & strUserDBType & "</td>" & vbNewLine & _
					"</tr>" & vbNewLine
	
	If strDBType = "sqlserver" then 
		Response.Write	"<tr>" & vbNewLine & _
						"<td class=""formlabel"">SQL Server Version:&nbsp;</td>" & vbNewLine & _
						"<td class=""formvalue"">" & vbNewLine & _
						"<input type=""radio"" name=""sqltype"" value=""7"" checked> SQL 7.x/2000&nbsp;(or greater)&nbsp;&nbsp;&nbsp;" & vbNewLine & _
						"<input type=""radio"" name=""sqltype"" value=""6""> SQL 6.x&nbsp;&nbsp;&nbsp;" & vbNewLine & _
						"</td>" & vbNewLine & _
						"</tr>" & vbNewLine
	end if
	Response.Write	"<tr>" & vbNewLine & _
					"<td class=""formlabel"">Code:&nbsp;</td>" & vbNewLine & _
					"<td class=""formvalue"">" & vbNewLine & _
					"<textarea name=""Message"" cols=""70"" rows=""15"" wrap=""VIRTUAL""></textarea></td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"<tr>" & vbNewLine & _
					"<td class=""options"" colspan=""2""><p>Enter the Code in the box above that you would like to process.<br />" & _
					"A script will execute to perform the database upgrade.</p>" & vbNewLine & _
					"<button name=""Submit"" type=""submit"">Submit</button>" & vbNewLine & _
					"</td>" & vbNewLine & _
					"</tr>" & vbNewLine & _
					"</table>" & vbNewLine & _
					"</form>" & vbNewLine
end if

WriteFooter
Response.End

Sub CreateTables(numfields)
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		TablePrefix = strMemberTablePrefix
	else
		TablePrefix = strTablePrefix
	end if

	strSql = "CREATE TABLE " & TablePrefix & strTableName & "( "
	if idFieldName <> "" then
		select case strDBType
			case "access"
				if Instr(strConnString,"(*.mdb)") then
					strSql = strSql & idFieldName &" COUNTER CONSTRAINT PrimaryKey PRIMARY KEY "
				else
					strSql = strSql & idFieldName &" int IDENTITY (1, 1) PRIMARY KEY NOT NULL "
				end if
			case "sqlserver"
				strSql = strSql & idFieldName &" int IDENTITY (1, 1) PRIMARY KEY NOT NULL "
			case "mysql"
				strSql = strSql & idFieldName &" INT (11) NOT NULL auto_increment "
		end select
	end if
	
	for y = 0 to numfields -1
	on error resume next
		tmpArray = split(fieldArray(y),"#")
		fName = uCase(tmpArray(0))
		fType = lCase(tmpArray(1))
		fNull = uCase(tmpArray(2))
		fDefault = tmpArray(3)
		
		if idFieldName <> "" or y <> 0 then
			strSql = strSql & ", "
		end if
		
		select case strDBType
			case "access"
				fType = replace(fType,"varchar (","text (")
			case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
						fType = replace(fType,"date","datetime")
					case else
						fType = replace(fType,"memo","text")
				end select
			case "mysql"
				fType = replace(fType,"memo","text")
				fType = replace(fType,"#int","#int (11)")
				fType = replace(fType,"#smallint","#smallint (6)")
		end select
		
		if fNull <> "NULL" then fNull = "NOT NULL"
		
		strSql = strSql & fName & " " & fType & " " & fNull & " " 
		
		if fdefault <> "" then
			select case strDBType
				case "access"
					if Instr(lcase(strConnString), "jet") then strSql = strSql & "DEFAULT " & fDefault
				case else
					strSql = strSql & "DEFAULT " & fDefault
			end select
		end if
	next
	
	if strDBType = "mysql" then
		if idFieldName <> "" then
			strSql = strSql & ",KEY " & TablePrefix & strTableName & "_" & idFieldName & "(" & idFieldName & "))"
		else
			strSql = strSql & ")"
		end if
	else
		strSql = strSql & ")"
	end if
	
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	if err.number <> 0 and err.number <> 13 and err.number <> tableExists then
		response.Write	"<div class=""oops modsetup"">" & vbNewLine & _
						"<p>" & strSql & "</p>" & vbNewLine & _
						"<p>Error #" & err.number & ": " & err.description & "</p>" & vbNewLine & _
						"</div>" & vbNewLine
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableExists then 
			response.Write	"<div class=""warning modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Table already exists.</p>" & vbNewLine & _
							"</div>" & vbNewLine
		else
			response.Write	"<div class=""ok modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Table created successfully.</p>" & vbNewLine & _
							"</div>" & vbNewLine
		end if
	end if
end Sub

Sub AlterTables(numfields)
	for y = 0 to numfields -1
		on error resume next
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			TablePrefix = strMemberTablePrefix
		else
			TablePrefix = strTablePrefix
		end if
		strSql = "ALTER TABLE " & TablePrefix & strTableName 
		tmpArray = split(fieldArray(y),"#")
		fAction = uCase(tmpArray(0))
		fName = uCase(tmpArray(1))
		fType = lCase(tmpArray(2))
		fNull = uCase(tmpArray(3))
		fDefault = tmpArray(4)
		select case fAction
			case "ADD"
				strSQL = strSQL & " ADD "
				if strDBType = "access" then strSql = strSql & "COLUMN "
			case "DROP"
				strSQL = strSQL & " DROP COLUMN "
			case "ALTER"
				strSQL = strSQL & " ALTER COLUMN "
			case else
		end select
		if fAction = "ADD" or fAction = "ALTER" then
			select case strDBType
				case "access"
					fType = replace(fType,"varchar (","text (")
				case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
						fType = replace(fType,"date","datetime")
					case else
						fType = replace(fType,"memo","text")
				end select
				case "mysql"
					fType = replace(fType,"memo","text")
					fType = replace(fType,"#int","#int (11)")
					fType = replace(fType,"#smallint","#smallint (6)")
			end select
			if fNull <> "NULL" then fNull = "NOT NULL"
			strSql = strSQL & fName & " " & fType & " " & fNULL & " "
			if fDefault <> "" then
				select case strDBType
					case "access"
						if Instr(lcase(strConnString), "jet") then strSql = strSql & "DEFAULT " & fDefault
					case else
						strSql = strSql & "DEFAULT " & fDefault
				end select
			end if
		else
			strSql = strSQL & fName
		end if
		
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		if err.number <> 0 and err.number <> 13 and err.number <> fieldExists then
			response.Write	"<div class=""oops modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Error #" & err.number & ": " & err.description & "</p>" & vbNewLine & _
							"</div>" & vbNewLine
			ErrorCount = ErrorCount + 1
			resultString = ""
		else
			if fAction = "DROP" then
				response.Write	"<div class=""ok modsetup"">" & vbNewLine & _
								"<p>" & strSql & "</p>" & vbNewLine & _
								"<p>Column " & LCase(fAction) & "ped successfully.</p>" & vbNewLine & _
								"</div>" & vbNewLine
			else
				if err.number = fieldExists then 
					response.Write	"<div class=""warning modsetup"">" & vbNewLine & _
									"<p>" & strSql & "</p>" & vbNewLine & _
									"<p>Column already exists.</p>" & vbNewLine & _
									"</div>" & vbNewLine
				else
					response.Write	"<div class=""ok modsetup"">" & vbNewLine & _
									"<p>" & strSql & "</p>" & vbNewLine & _
									"<p>Column " & LCase(fAction) & "ed successfully.</p>" & vbNewLine & _
									"</div>" & vbNewLine
				end if
			end if
			if fDefault <> "" and err.number <> fieldExists then
				strSQL = "UPDATE " & TablePrefix & strTableName & " SET " & fName & "=" & fDefault
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				response.write	"<div class=""ok modsetup"">" & vbNewLine & _
								"<p>" & strSql & "</p>" & vbNewLine & _
								"<p>Populating current records with new default value.</p>" & vbNewLine & _
								"</div>" & vbNewLine
			end if
		end if

		if fieldArray(y) = "" then y = numfields
	next
end Sub

Sub InsertValues(numfields)
	on error resume next
	for y = 0 to numfields-1
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			strSql = "INSERT INTO " & strMemberTablePrefix & strTableName & " "
		else
			strSql = "INSERT INTO " & strTablePrefix & strTableName & " "
		end if
		tmpArray = split(fieldArray(y),"#")
		fNames = tmpArray(0)
		fValues = tmpArray(1)
		strSql = strSql & tmpArray(0) & " VALUES " & tmpArray(1)
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 and err.number <> 13 then
			response.Write	"<div class=""oops modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Error #" & err.number & ": " & err.description & "</p>" & vbNewLine & _
							"</div>" & vbNewLine
		ErrorCount = ErrorCount + 1
	else
			response.write	"<div class=""ok modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Value(s) updated successfully.</p>" & vbNewLine & _
							"</div>" & vbNewLine
	end if
	next
end Sub 

Sub UpdateValues(numfields)
	on error resume next
	for y = 0 to numfields-1
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			strSql = "UPDATE " & strMemberTablePrefix & strTableName & " SET"
		else
			strSql = "UPDATE " & strTablePrefix & strTableName & " SET"
		end if
		tmpArray = split(fieldArray(y),"#")
		fName = tmpArray(0)
		fValue = tmpArray(1)
		fWhere = tmpArray(2)
		strSql = strSql & " " & fName & " = " & fvalue
		if fWhere <> "" then
			strSql = strSql & " WHERE " & fWhere
		end if
		
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		if err.number <> 0 and err.number <> 13 then
			response.Write	"<div class=""oops modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Error #" & err.number & ": " & err.description & "</p>" & vbNewLine & _
							"</div>" & vbNewLine
		ErrorCount = ErrorCount + 1
	else
			response.write	"<div class=""ok modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Value(s) updated successfully.</p>" & vbNewLine & _
							"</div>" & vbNewLine
	end if
	next
end Sub 

Sub DeleteValues(numfields)
	on error resume next
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		strSql = "DELETE FROM " & strMemberTablePrefix & strTableName & " WHERE "
	else
		strSql = "DELETE FROM " & strTablePrefix & strTableName & " WHERE "
	end if
	tmpArray = fieldArray(0)
	strSql = strSql & tmpArray
	
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 then
		response.Write	"<div class=""oops modsetup"">" & vbNewLine & _
						"<p>" & strSql & "</p>" & vbNewLine & _
						"<p>Error #" & err.number & ": " & err.description & "</p>" & vbNewLine & _
						"</div>" & vbNewLine
		ErrorCount = ErrorCount + 1
	else
		response.write	"<div class=""ok modsetup"">" & vbNewLine & _
						"<p>" & strSql & "</p>" & vbNewLine & _
						"<p>Value(s) updated successfully.</p>" & vbNewLine & _
						"</div>" & vbNewLine
	end if
end Sub 

Sub DropTable()
	on error resume next
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		strSql = "DROP TABLE " & strMemberTablePrefix & strTableName
	else
		strSql = "DROP TABLE " & strTablePrefix & strTableName
	end if
	
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 and err.number <> 13 and err.number <> tableNotExist then
		response.Write	"<div class=""oops modsetup"">" & vbNewLine & _
						"<p>" & strSql & "</p>" & vbNewLine & _
						"<p>Error #" & err.number & ": " & err.description & "</p>" & vbNewLine & _
						"</div>" & vbNewLine
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableNotExist then
			response.Write	"<div class=""warning modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Table does not exist.</p>" & vbNewLine & _
							"</div>" & vbNewLine
		else
			response.write	"<div class=""ok modsetup"">" & vbNewLine & _
							"<p>" & strSql & "</p>" & vbNewLine & _
							"<p>Table dropped successfully.</p>" & vbNewLine & _
							"</div>" & vbNewLine
		end if
	end if
end Sub

on error goto 0
%>
