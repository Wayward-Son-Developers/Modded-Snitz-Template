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

Response.Write	"<option value=""Afghanistan"">Afghanistan</option>" & vbNewLine & _
				"<option value=""Albania"">Albania</option>" & vbNewLine & _
				"<option value=""Algeria"">Algeria</option>" & vbNewLine & _
				"<option value=""Andorra"">Andorra</option>" & vbNewLine & _
				"<option value=""Angola"">Angola</option>" & vbNewLine & _
				"<option value=""Anguilla"">Anguilla</option>" & vbNewLine & _
				"<option value=""Antigua and Barbuda"">Antigua and Barbuda</option>" & vbNewLine & _
				"<option value=""Argentina"">Argentina</option>" & vbNewLine & _
				"<option value=""Armenia"">Armenia</option>" & vbNewLine & _
				"<option value=""Aruba"">Aruba</option>" & vbNewLine & _
				"<option value=""Australia"">Australia</option>" & vbNewLine & _
				"<option value=""Austria"">Austria</option>" & vbNewLine & _
				"<option value=""Azerbaijan"">Azerbaijan</option>" & vbNewLine & _
				"<option value=""Azores"">Azores</option>" & vbNewLine & _
				"<option value=""Bahamas"">Bahamas</option>" & vbNewLine & _
				"<option value=""Bahrain"">Bahrain</option>" & vbNewLine & _
				"<option value=""Bangladesh"">Bangladesh</option>" & vbNewLine & _
				"<option value=""Barbados"">Barbados</option>" & vbNewLine & _
				"<option value=""Belarus"">Belarus</option>" & vbNewLine & _
				"<option value=""Belgium"">Belgium</option>" & vbNewLine & _
				"<option value=""Belize"">Belize</option>" & vbNewLine & _
				"<option value=""Benin"">Benin</option>" & vbNewLine & _
				"<option value=""Bermuda"">Bermuda</option>" & vbNewLine & _
				"<option value=""Bhutan"">Bhutan</option>" & vbNewLine & _
				"<option value=""Bolivia"">Bolivia</option>" & vbNewLine & _
				"<option value=""Borneo"">Borneo</option>" & vbNewLine & _
				"<option value=""Bosnia and Herzegovina"">Bosnia and Herzegovina</option>" & vbNewLine & _
				"<option value=""Botswana"">Botswana</option>" & vbNewLine & _
				"<option value=""Brazil"">Brazil</option>" & vbNewLine & _
				"<option value=""British Indian Ocean Territories"">British Indian Ocean Territories</option>" & vbNewLine & _
				"<option value=""Brunei"">Brunei</option>" & vbNewLine & _
				"<option value=""Bulgaria"">Bulgaria</option>" & vbNewLine & _
				"<option value=""Burkina Faso (Upper Volta)"">Burkina Faso (Upper Volta)</option>" & vbNewLine & _
				"<option value=""Burundi"">Burundi</option>" & vbNewLine & _
				"<option value=""Camaroon"">Camaroon</option>" & vbNewLine & _
				"<option value=""Cambodia"">Cambodia</option>" & vbNewLine & _
				"<option value=""Canada"">Canada</option>" & vbNewLine & _
				"<option value=""Canary Islands"">Canary Islands</option>" & vbNewLine & _
				"<option value=""Cape Vere Islands"">Cape Vere Islands</option>" & vbNewLine & _
				"<option value=""Cayman Islands"">Cayman Islands</option>" & vbNewLine & _
				"<option value=""Central African Rep"">Central African Rep</option>" & vbNewLine & _
				"<option value=""Chad"">Chad</option>" & vbNewLine & _
				"<option value=""Chile"">Chile</option>" & vbNewLine & _
				"<option value=""China"">China</option>" & vbNewLine & _
				"<option value=""Christmas Island"">Christmas Island</option>" & vbNewLine & _
				"<option value=""Colombia"">Colombia</option>" & vbNewLine & _
				"<option value=""Comoros Islands"">Comoros Islands</option>" & vbNewLine & _
				"<option value=""Congo, Democratic Republic of"">Congo, Democratic Republic of</option>" & vbNewLine & _
				"<option value=""Costa Rica"">Costa Rica</option>" & vbNewLine & _
				"<option value=""Croatia"">Croatia</option>" & vbNewLine & _
				"<option value=""Cuba"">Cuba</option>" & vbNewLine & _
				"<option value=""Cyprus"">Cyprus</option>" & vbNewLine & _
				"<option value=""Czech Republic"">Czech Republic</option>" & vbNewLine & _
				"<option value=""Denmark"">Denmark</option>" & vbNewLine & _
				"<option value=""Djibouti"">Djibouti</option>" & vbNewLine & _
				"<option value=""Dominica"">Dominica</option>" & vbNewLine & _
				"<option value=""Dominican Republic"">Dominican Republic</option>" & vbNewLine & _
				"<option value=""East Timor"">East Timor</option>" & vbNewLine & _
				"<option value=""Ecuador"">Ecuador</option>" & vbNewLine & _
				"<option value=""Egypt"">Egypt</option>" & vbNewLine & _
				"<option value=""El Salvador"">El Salvador</option>" & vbNewLine & _
				"<option value=""Equatorial Guinea"">Equatorial Guinea</option>" & vbNewLine & _
				"<option value=""Eritria"">Eritria</option>" & vbNewLine & _
				"<option value=""Estonia"">Estonia</option>" & vbNewLine & _
				"<option value=""Ethiopia"">Ethiopia</option>" & vbNewLine & _
				"<option value=""Falkland Islands"">Falkland Islands</option>" & vbNewLine & _
				"<option value=""Faroe Islands"">Faroe Islands</option>" & vbNewLine & _
				"<option value=""Fed Rep Yugoslavia"">Fed Rep Yugoslavia</option>" & vbNewLine & _
				"<option value=""Fiji"">Fiji</option>" & vbNewLine & _
				"<option value=""Finland"">Finland</option>" & vbNewLine & _
				"<option value=""France"">France</option>" & vbNewLine & _
				"<option value=""French Guiana"">French Guiana</option>" & vbNewLine & _
				"<option value=""French Polynesia"">French Polynesia</option>" & vbNewLine & _
				"<option value=""Fyro Macedonia"">Fyro Macedonia</option>" & vbNewLine & _
				"<option value=""Gabon"">Gabon</option>" & vbNewLine & _
				"<option value=""Gambia"">Gambia</option>" & vbNewLine & _
				"<option value=""Georgia"">Georgia</option>" & vbNewLine & _
				"<option value=""Germany"">Germany</option>" & vbNewLine & _
				"<option value=""Ghana"">Ghana</option>" & vbNewLine & _
				"<option value=""Gibraltar"">Gibraltar</option>" & vbNewLine & _
				"<option value=""Greece"">Greece</option>" & vbNewLine & _
				"<option value=""Greenland"">Greenland</option>" & vbNewLine & _
				"<option value=""Grenada"">Grenada</option>" & vbNewLine & _
				"<option value=""Guadeloupe"">Guadeloupe</option>" & vbNewLine & _
				"<option value=""Guatemala"">Guatemala</option>" & vbNewLine & _
				"<option value=""Guinea"">Guinea</option>" & vbNewLine & _
				"<option value=""Guinea-Bissau"">Guinea-Bissau</option>" & vbNewLine & _
				"<option value=""Guyana"">Guyana</option>" & vbNewLine & _
				"<option value=""Haiti"">Haiti</option>" & vbNewLine & _
				"<option value=""Honduras"">Honduras</option>" & vbNewLine & _
				"<option value=""Hong Kong"">Hong Kong</option>" & vbNewLine & _
				"<option value=""Hungary"">Hungary</option>" & vbNewLine & _
				"<option value=""Iceland"">Iceland</option>" & vbNewLine & _
				"<option value=""India"">India</option>" & vbNewLine & _
				"<option value=""Indonesia"">Indonesia</option>" & vbNewLine & _
				"<option value=""Iran"">Iran</option>" & vbNewLine & _
				"<option value=""Iraq"">Iraq</option>" & vbNewLine & _
				"<option value=""Ireland"">Ireland</option>" & vbNewLine & _
				"<option value=""Israel"">Israel</option>" & vbNewLine & _
				"<option value=""Italy"">Italy</option>" & vbNewLine & _
				"<option value=""Ivory Coast"">Ivory Coast</option>" & vbNewLine & _
				"<option value=""Jamaica"">Jamaica</option>" & vbNewLine & _
				"<option value=""Japan"">Japan</option>" & vbNewLine & _
				"<option value=""Jordan"">Jordan</option>" & vbNewLine & _
				"<option value=""Kazakhstan"">Kazakhstan</option>" & vbNewLine & _
				"<option value=""Kenya"">Kenya</option>" & vbNewLine & _
				"<option value=""Kiribati"">Kiribati</option>" & vbNewLine & _
				"<option value=""Korea"">Korea</option>" & vbNewLine & _
				"<option value=""Kuwait"">Kuwait</option>" & vbNewLine & _
				"<option value=""Kyrgyzstan"">Kyrgyzstan</option>" & vbNewLine & _
				"<option value=""Laos"">Laos</option>" & vbNewLine & _
				"<option value=""Latvia"">Latvia</option>" & vbNewLine & _
				"<option value=""Lebanon"">Lebanon</option>" & vbNewLine & _
				"<option value=""Lesotho"">Lesotho</option>" & vbNewLine & _
				"<option value=""Liberia"">Liberia</option>" & vbNewLine & _
				"<option value=""Libya"">Libya</option>" & vbNewLine & _
				"<option value=""Liechtenstein"">Liechtenstein</option>" & vbNewLine & _
				"<option value=""Lithuania"">Lithuania</option>" & vbNewLine & _
				"<option value=""Luxembourg"">Luxembourg</option>" & vbNewLine & _
				"<option value=""Macao"">Macao</option>" & vbNewLine & _
				"<option value=""Madagascar"">Madagascar</option>" & vbNewLine & _
				"<option value=""Malawi"">Malawi</option>" & vbNewLine & _
				"<option value=""Malaysia"">Malaysia</option>" & vbNewLine & _
				"<option value=""Maldives"">Maldives</option>" & vbNewLine & _
				"<option value=""Mali"">Mali</option>" & vbNewLine & _
				"<option value=""Malta"">Malta</option>" & vbNewLine & _
				"<option value=""Martinique"">Martinique</option>" & vbNewLine & _
				"<option value=""Mauritania"">Mauritania</option>" & vbNewLine & _
				"<option value=""Mauritius"">Mauritius</option>" & vbNewLine & _
				"<option value=""Mexico"">Mexico</option>" & vbNewLine & _
				"<option value=""Moldova"">Moldova</option>" & vbNewLine & _
				"<option value=""Monaco"">Monaco</option>" & vbNewLine & _
				"<option value=""Mongolia"">Mongolia</option>" & vbNewLine & _
				"<option value=""Montserrat"">Montserrat</option>" & vbNewLine & _
				"<option value=""Morocco"">Morocco</option>" & vbNewLine & _
				"<option value=""Mozambique"">Mozambique</option>" & vbNewLine & _
				"<option value=""Myanmar (Burma)"">Myanmar (Burma)</option>" & vbNewLine & _
				"<option value=""Namibia"">Namibia</option>" & vbNewLine & _
				"<option value=""Naura"">Naura</option>" & vbNewLine & _
				"<option value=""Nepal"">Nepal</option>" & vbNewLine & _
				"<option value=""Netherlands"">Netherlands</option>" & vbNewLine & _
				"<option value=""Netherlands Antilles"">Netherlands Antilles</option>" & vbNewLine & _
				"<option value=""New Caledonia"">New Caledonia</option>" & vbNewLine & _
				"<option value=""New Zealand"">New Zealand</option>" & vbNewLine & _
				"<option value=""Nicaragua"">Nicaragua</option>" & vbNewLine & _
				"<option value=""Niger"">Niger</option>" & vbNewLine & _
				"<option value=""Nigeria"">Nigeria</option>" & vbNewLine & _
				"<option value=""Niue"">Niue</option>" & vbNewLine & _
				"<option value=""Norway"">Norway</option>" & vbNewLine & _
				"<option value=""Oman"">Oman</option>" & vbNewLine & _
				"<option value=""Pakistan"">Pakistan</option>" & vbNewLine & _
				"<option value=""Panama"">Panama</option>" & vbNewLine & _
				"<option value=""Papua New Guinea"">Papua New Guinea</option>" & vbNewLine & _
				"<option value=""Paraguay"">Paraguay</option>" & vbNewLine & _
				"<option value=""Peru"">Peru</option>" & vbNewLine & _
				"<option value=""Philippines"">Philippines</option>" & vbNewLine & _
				"<option value=""Pitcairn Island"">Pitcairn Island</option>" & vbNewLine & _
				"<option value=""Poland"">Poland</option>" & vbNewLine & _
				"<option value=""Portugal"">Portugal</option>" & vbNewLine & _
				"<option value=""Qatar"">Qatar</option>" & vbNewLine & _
				"<option value=""Republic of Korea"">Republic of Korea</option>" & vbNewLine & _
				"<option value=""Reunion Island"">Reunion Island</option>" & vbNewLine & _
				"<option value=""Romania"">Romania</option>" & vbNewLine & _
				"<option value=""Russia"">Russia</option>" & vbNewLine & _
				"<option value=""Rwanda"">Rwanda</option>" & vbNewLine & _
				"<option value=""Saint Barthelemy"">Saint Barthelemy</option>" & vbNewLine & _
				"<option value=""Saint Croix"">Saint Croix</option>" & vbNewLine & _
				"<option value=""Saint Helena"">Saint Helena</option>" & vbNewLine & _
				"<option value=""Saint Kitts and Nevis"">Saint Kitts and Nevis</option>" & vbNewLine & _
				"<option value=""Saint Lucia"">Saint Lucia</option>" & vbNewLine & _
				"<option value=""Saint Pierre and Miquelon"">Saint Pierre and Miquelon</option>" & vbNewLine & _
				"<option value=""Saint Vincent and Grenadi"">Saint Vincent and Grenadi</option>" & vbNewLine & _
				"<option value=""San Marino"">San Marino</option>" & vbNewLine & _
				"<option value=""Sao Tome and Principe"">Sao Tome and Principe</option>" & vbNewLine & _
				"<option value=""Saudi Arabia"">Saudi Arabia</option>" & vbNewLine & _
				"<option value=""Senegal"">Senegal</option>" & vbNewLine & _
				"<option value=""Seychelles"">Seychelles</option>" & vbNewLine & _
				"<option value=""Sierra Leone"">Sierra Leone</option>" & vbNewLine & _
				"<option value=""Singapore"">Singapore</option>" & vbNewLine & _
				"<option value=""Slovakia"">Slovakia</option>" & vbNewLine & _
				"<option value=""Slovenia"">Slovenia</option>" & vbNewLine & _
				"<option value=""Solomon Islands"">Solomon Islands</option>" & vbNewLine & _
				"<option value=""Somalia Northern Region"">Somalia Northern Region</option>" & vbNewLine & _
				"<option value=""Somalia Southern Region"">Somalia Southern Region</option>" & vbNewLine & _
				"<option value=""South Africa"">South Africa</option>" & vbNewLine & _
				"<option value=""South Sandwich Islands"">South Sandwich Islands</option>" & vbNewLine & _
				"<option value=""Spain"">Spain</option>" & vbNewLine & _
				"<option value=""Sri Lanka"">Sri Lanka</option>" & vbNewLine & _
				"<option value=""Sudan"">Sudan</option>" & vbNewLine & _
				"<option value=""Suriname"">Suriname</option>" & vbNewLine & _
				"<option value=""Swaziland"">Swaziland</option>" & vbNewLine & _
				"<option value=""Sweden"">Sweden</option>" & vbNewLine & _
				"<option value=""Switzerland"">Switzerland</option>" & vbNewLine & _
				"<option value=""Syria"">Syria</option>" & vbNewLine & _
				"<option value=""Taiwan"">Taiwan</option>" & vbNewLine & _
				"<option value=""Tajikistan"">Tajikistan</option>" & vbNewLine & _
				"<option value=""Tanzania"">Tanzania</option>" & vbNewLine & _
				"<option value=""Thailand"">Thailand</option>" & vbNewLine & _
				"<option value=""Togo"">Togo</option>" & vbNewLine & _
				"<option value=""Tonga"">Tonga</option>" & vbNewLine & _
				"<option value=""Trinidad and Tobago"">Trinidad and Tobago</option>" & vbNewLine & _
				"<option value=""Tunisia"">Tunisia</option>" & vbNewLine & _
				"<option value=""Turkey"">Turkey</option>" & vbNewLine & _
				"<option value=""Turkmenistan"">Turkmenistan</option>" & vbNewLine & _
				"<option value=""Turks and Caicos Islnd"">Turks and Caicos Islnd</option>" & vbNewLine & _
				"<option value=""Tuvalu"">Tuvalu</option>" & vbNewLine & _
				"<option value=""USA"">USA</option>" & vbNewLine & _
				"<option value=""Uganda"">Uganda</option>" & vbNewLine & _
				"<option value=""Ukraine"">Ukraine</option>" & vbNewLine & _
				"<option value=""United Arab Emirates"">United Arab Emirates</option>" & vbNewLine & _
				"<option value=""United Kingdom"">United Kingdom</option>" & vbNewLine & _
				"<option value=""Uruguay"">Uruguay</option>" & vbNewLine & _
				"<option value=""Uzbekistan"">Uzbekistan</option>" & vbNewLine & _
				"<option value=""Vanuatu"">Vanuatu</option>" & vbNewLine & _
				"<option value=""Vatican City"">Vatican City</option>" & vbNewLine & _
				"<option value=""Venezuela"">Venezuela</option>" & vbNewLine & _
				"<option value=""Vietnam"">Vietnam</option>" & vbNewLine & _
				"<option value=""Virgin Islands (United Kingdom)"">Virgin Islands (United Kingdom)</option>" & vbNewLine & _
				"<option value=""Wallis and Futuna Islands"">Wallis and Futuna Islands</option>" & vbNewLine & _
				"<option value=""Western Sahara"">Western Sahara</option>" & vbNewLine & _
				"<option value=""Western Samoa"">Western Samoa</option>" & vbNewLine & _
				"<option value=""Yemen"">Yemen</option>" & vbNewLine & _
				"<option value=""Yugoslavia"">Yugoslavia</option>" & vbNewLine & _
				"<option value=""Zambia"">Zambia</option>" & vbNewLine & _
				"<option value=""Zimbabwe"">Zimbabwe</option>" & vbNewLine
%>
