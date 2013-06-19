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

function to_base (iNumber,base)
		base_string = "0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
		base_chars = split(base_string,",")
        if (iNumber=0) then
                result =  0
		end if
		
        result = ""
        while (iNumber > 0)
			ii = iNumber mod base
			result = base_chars(CInt(ii)) & result
            iNumber = (iNumber- ii) \ base
        wend
		
        to_base = result
end function

function from_base (iString,base)
        base_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
		result = 0
        for x = 1 to len(iString)
			result = result * base
			c = Mid(iString,x,1)

			pos = Instr(base_chars,c)-1
			result = result + pos
		next
		from_base = result
end function

Function RandomNumber(intHighestNumber)
	if intHighestNumber > 60 then intHighestNumber = 60
	Randomize
	RandomNumber = Int(intHighestNumber * Rnd) + 2
End Function

	'###################### Random Field checker ########################################
Dim randSeed
Dim encFieldname
Dim myDate

  myDate = doublenum(Month(Now)) & doublenum(Day(Now)) & doublenum(Hour(Now)) & doublenum(Minute(Now)) & doublenum(second(Now))
  if Session("FormChecker") = "" then
  		Session("FormChecker") = myDate
		Session("randSeed") = RandomNumber(60)
  end if

encFieldname = to_base(Session("FormChecker"),Session("randSeed"))
	'###################### End Random Field checker #####################################
%>
