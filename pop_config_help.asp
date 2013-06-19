<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'## Huw Reddick and Richard Kinser
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
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA02111-1307, USA.
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
Response.Write	"<table class=""content"" width=""100%"">" & vbNewLine
select case Request.QueryString("mode")
	case "system"
		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""strConnString""></a>How do I configure the strConnString?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td><ul>" & vbNewLine & _
						"<li><b>DSN:</b><br />" & vbNewLine & _
						"snitz_forum</li>" & vbNewLine & _
						"<li><b>MS Access DSN-less:</b><br />" & vbNewLine & _
						"strConnString = &quot;DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\www\snitz.com\db\snitz_forum.mdb&quot;</li>" & vbNewLine & _
						"</ul><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""tableprefix""></a>What's Table Name Prefix?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Table Name Prefix is used if you have multiple versions of the forum running in the same database. This way you can name the tables differently and still use one user to connect. (eg. FORUM_ and FORUM2_)" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""forumtitle""></a>What's Forum Title?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Forum Title is the title that shows up in the upper right hand corner of the forum. It is also used in e-mails to show where the e-mail came from when posting replies are sent and when new users register." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""copyright""></a>What's Forum Copyright?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This copyright statements location is basically saying that any topics or replies that are posted are copyrighted material of your organization. This copyright location also helps to copyright the images of your logo and any other material that may be posted on forum pages; however, it is understood by copyright statements in code and informational pages, that the forum code itself is still copyright &copy; 2000 Snitz Communications.<br /><br />" & _
						"<span class=""hlf""><b>NOTE:</b>The &copy; will be included automatically.</span>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""titleimage""></a>What's Title Image?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Use a relative path to point to the image you want to show up in the upper left-hand corner of your forum window.<br />" & vbNewLine & _
						"<br />" & vbNewLine & _
						"For example:<br />" & vbNewLine & _
						"<b>bboard_snitz.gif</b><br />" & vbNewLine & _
						"This points to the bboard_snitz.gif graphic in the same directory, whereas the following would point to the root of the web server and up into the base /images/ directory:<br />" & vbNewLine & _
						"<b>../images/bboard_snitz.gif</b>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""homeurl""></a>What's the Home URL?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"The Home URL is the base address for your website. An example would be:<br />" & vbNewLine & _
						"<b>forum.snitz.com</b><br />" & vbNewLine & _
						"<br />" & vbNewLine & _
						"<span class=""hlf"">NOTE: Include the full path of the URL whether it begins with <b>http://</b> in front or a relative URL such as <b>../</b>.</span>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""forumurl""></a>What's the Forum URL?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"The Forum URL is the base address for your forum. An example would be:<br />" & vbNewLine & _
						"<b>http://forum.snitz.com/forum</b>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""imagelocation""></a>What is the Images Location?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Enter the location where your images are located.<br />" & vbNewLine & _
						"If you have not moved the images from their default location, then just leave this field blank.<br /><br />" & vbNewLine & _
						"But, if you have created an <b>images</b> directory in your <b>forum</b> directory then enter:<br /><br />" & vbNewLine & _
						"<b>images/</b><br /><br />in the field." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""AuthType""></a>Authorization Type?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"You can either select DataBase or NT Domain authorization." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""SetCookieToForum""></a>Set Cookie To...</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"You can tell your forum to set it's cookie to either the forum, or the base website. You would set it to the forum if you were hosting multiple forums on the same server or the same domain, and they each had different user communities, otherwise you want this feature set to Website and NOT Forum." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""GfxButtons""></a>Graphic Buttons?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"By enabling this feature, the forums will use pictures/graphics instead of the default buttons for ""Submit"" and ""Reset"" etc..." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""PoweredBy""></a>Use Graphic for ""Powered By"" link?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Toggles between using a Graphic Powered By Link, or a Text Powered By Link.Either way, you must have one or the other..." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ProhibitNewMembers""></a>Prohibit New Members?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Toggles between allowing or disallowing people to Register on your Forum." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""RequireReg""></a>Require Registration?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"When this option in set to <b>On</b>, only registered members who are logged in will be able to view your Forum.Everyone else will be presented with a login screen." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""UserNameFilter""></a>UserName Filter?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"When this option in set to <b>On</b>, the names (or names that contain words) that you specify in the UserName Filter configuration will not be available for user's to register with." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	case "features"
		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""secureadminmode""></a>Secure Admin Mode?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td><span class=""hlf"">" & vbNewLine & _
						"<b>WARNING: Only turn Secure Admin off if you absolutely need to. If this option is turned off, anyone can change your forum's configuration!</b></span>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""allownoncookies""></a>Why would I want Non-Cookie Mode on?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"If your user base does not use cookies, then you would want to turn this function ""ON"". WARNING: all your admin functions will be visible to all users if this function is ""ON""." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""IPLogging""></a>What is IP Logging?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"IP Logging will record in the database the IP address of the person who posted a new Topic or Reply. A moderator or administrator then could view the IP by clicking on an icon above the post in the topic." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""FloodCheck""></a>What is Flood Control?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"With Flood Control enabled, normal users will have to wait the specified amount of time between posts before they can post again." & vbNewLine & _
						"<br /><br />Admins and Moderators are not affected by this limitation." & vbNewLine & _
						"<br /><br />You can choose 30 seconds, 60 seconds, 90 seconds or 120 seconds." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""privateforums""></a>What are Private Forums?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Private Forums enable you to only allow certain members to see that the forum exists. If it's only password protected, everyone can see that it exists, however, they are prompted for a password to get in." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""groupcategories""></a>What are Group Categories?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Group Categories enable you to ""group"" Categories together into ""Groups"" to better organize how Categories are displayed on your forum." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Subscription""></a>What is Highest level of Subscription for?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allows you to set the Highest Level of Subscription that can be used on the Forum.You will also need to set the individual level in each of your Categories and Forums." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""badwordfilter""></a>Bad Word Filter?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Screen out words you and your guests would find offensive.<br /><br />Bad Words can by configured via the Bad Word Configuration option in the Admin Options." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Moderation""></a>What does Allow Topic Moderation do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"When enabled, this feature allows the Administrator or the Moderator to ""Approve"", ""Hold"" or ""Delete"" a users post before it is shown to the public." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowModerator""></a>What does Show Moderators do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Basically, if this function is on, it shows the name of the moderator beside the forum that they moderate on the main default page. If it is off, however, visitors won't see who is moderating the forum they are posting in." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""MoveTopicMode""></a>Why Restrict Moderators from Moving Posts?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This feature either allows or dis-allows a Moderator of one forum to move topics within their forum to someone else's forum where they do not have moderator rights." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""MoveNotify""></a>Can I notify the Author if his Topic is moved?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"If enabled, this feature automatically sends an e-mail to the topic author if it is moved." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ArchiveState""></a>What are Archive Functions?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This toggles whether the icons/links show up for the Archive Functions of this Forum." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""stats""></a>What does Show Detailed Statistics do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the display of detailed statistics (last visited date and time, last post, active topics, newest member) at the bottom of the forum." & vbNewLine & _
						"When turned off, some statistics are displayed at the top of the page." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""JumpLastPost""></a>What does Show Jump To Last Post Link do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the display of a Jump To Last Post Link " & getCurrentIcon(strIconLastpost,"","align=""absmiddle""") & " icon on the Default page, Forum page and Active Topics page.This link will take the user to the last post in that topic." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""showpaging""></a>What does Show Quick Paging do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Quick Paging is when you have a topic that is more than 1 page, a small graphic and the #'s will be show next to the topic title so you can go straight to page 2 or 3, etc..." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""pagenumbersize""></a>What is Pagenumbers per row for?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This is now only used for the Topic Paging, it limits the amount of pages shown in each row when a topic is more than one page long." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""StickyTopic""></a>What does Allow Sticky Topics do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the ability of an Admin or Moderator to ""Stick"" a post at the top of the Topics List.While this Topic is ""Sticky"", it will remain at the top of the list." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""editedbydate""></a>What would Edited By on Date do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"When a post is edited, there is an appending to the end of the post that says when and by whom the post was edited. Turning this function off would make it so that the footer would not be placed on the end of the post." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowTopicNav""></a>What does Show Prev / Next Topic do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the display of previous topic " & getCurrentIcon(strIconGoLeft,"","align=""absmiddle""") & " and next topic " & getCurrentIcon(strIconGoRight,"","align=""absmiddle""") & " icons on the topics page." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowSendToFriend""></a>What does Show Send to a Friend Link do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the display of a Send Topic to a Friend Link that is shown when viewing a topic..This link will allow a user to e-mail a topic to a friend.E-mail functions must be on for this link to show up." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowPrinterFriendly""></a>What does Show Printer Friendly Link do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the display of a Printer Friendly link that is shown when viewing a topic.This link will popup a window with the topic and any replies that are shown in a format that is easier to print." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""hottopics""></a>What are Hot Topics?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Hot Topics change the topic folder icon in the Forum view from a normal folder to a flaming folder to let people know that your minimum number of posts has been met to categorize this topic as one that is seeing a lot of action." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""pagesize""></a>What is Items per page for?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This is the maximum amount of items shown on each page. Once the amount of items on the page reaches this amount, a dropdown box will be shown where you can select other pages." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""AllowHTML""></a>Why would I allow HTML?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"By allowing HTML you are opening up a whole big can of worms. You may wish to allow HTML in a controlled INTRANET environment,though. It is not recommended to be used on the INTERNET as anyone can post anything without your being able to screen it. IE Pornographic pictures, JavaScript that messes up your pages, etc..." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""AllowForumCode""></a>Enable Forum Code?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"By turning off Forum Code, you can allow users to mark up their posts with safe codes." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""imginposts""></a>Why enable Images in Posts?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allows users to place images into their Posts. However, you should be aware that this feature would allow anyone to post ANY image in your forums. This may lead to broken links and potentially objectionable material being displayed." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""icons""></a>What do Icons do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow users to post smiley faces and other icons allowed by the Forums within the body of their posts!" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""signatures""></a>Why enable Signatures?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allows users to set a ""Signature"" into their Posts. The same concerns mentioned for Images in Posts applies here as well." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""dsignatures""></a>Why enable Dynamic Signatures?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"First, you must have Signatures enabled to use Dynamic Signatures.With Dynamic Signatures enabled, the users signature is not added to the post until it is viewed, so if a person changes their signature, that change will apply to all posts made by that user.But, this will only apply to posts made while Dyanmic Signatures are enabled.Any signature that is already in a post won't be updated." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowFormatButtons""></a>Show Format Buttons?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This turns off or on the Format Section on the screen where your users post new topics/reply to existing topics.<br /><br /><span class=""hlf"">Note:&nbsp;You must also have Forum Code enabled on your forum to use this feature.</span>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowSmiliesTable""></a>Show Smilies Table?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allows users to insert smilies in their posts by clicking on the smilie in a small table shown to them in the post screen.<br /><br /><span class=""hlf"">Note:&nbsp;You must also have Icons enabled on your forum to use this feature.</span>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowQuickReply""></a>Show Quick Reply?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allows users to reply to a topic via a reply box at the bottom of the page when viewing a topic." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""timer""></a>What does Show Timer do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off the display of the time it took (in seconds) to generate/display the current page.This time is shown in the footer of every (non popup) page." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""timerphrase""></a>What is Timer Phrase?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This is what will display in the footer of every (non popup) page.The phrase must contain the <b>[TIMER]</b> placeholder.This is where the actual time will be in the phrase (it's dynamically inserted when the page is created)." & vbNewLine & _
						"<br /><br /><span class=""hlf""><b>Show Timer must be enabled for this to be used.</b></span>" & vbNewLine & _
						"<br /><br />The default is:<b>This page was generated in [TIMER] seconds.</b>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	case "members"
		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""FullName""></a>What is Fullname For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Full Name (First Name and Last Name), to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Picture""></a>What is Picture For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter a link to a Picture of themselves, to be viewed in their profile.<br /><br />As Admin, you should review the picture in each user's profile from time to time to be sure that the Picture linked to is appropriate for your Forum." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""RecentTopics""></a>What is Recent Topics For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"When Recent Topics is enabled, a list of the last 10 Topics posted to by a user will be shown in their Profile.<br /><br />This includes New Topics and replies to existing topics." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Sex""></a>What is Sex For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Sex (either Male or Female), to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Age""></a>What is Age For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their age, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""AgeDOB""></a>What is Birth Date For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Birth Date, from which their Age will be calculated and displayed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""MinAge""></a>What is Minimum Age for?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Prevent users under the age you specify here from registering. The default is 13 for COPPA compliancy but you can change it to anything you want. To turn this feature off completely, set the minimum age to 0." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""City""></a>What is City For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their City, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""State""></a>What is State For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their State, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Country""></a>What is Country For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to choose their Country, to be viewed in their profile and in each Topic or Reply they post." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""aim""></a>What is the AIM Option For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off features that allow users to enter their AIM username... then for other users to send them messages and/or add them to their buddy list." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""icq""></a>What is the ICQ Option For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off features that allow users to enter their ICQ number... then for other users to send them ICQ messages and/or see if they are online." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""msn""></a>What is the MSN Option For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off features that allow users to enter their MSN username... then for other users to view their MSN Username." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""yahoo""></a>What is the YAHOO Option For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Turns On/Off features that allow users to enter their YAHOO username... then for other users to send them messages and/or add them to their buddy list." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Occupation""></a>What is Occupation For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Occupation, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Homepages""></a>What is Homepages For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to display their homepage link by their name on each post and in their Profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""FavLinks""></a>What is Favorite Links For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter 2 of their Favorite Links, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""MStatus""></a>What is Marital Status For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Marital Status, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Bio""></a>What is Bio For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Bio, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Hobbies""></a>What is Hobbies For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Hobbies, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""LNews""></a>What is Latest News For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Latest News, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""Quote""></a>What is Quote For?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Allow your users to enter their Quote, to be viewed in their profile." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	case "ranks"
		arrStarColors = ("Gold|Silver|Bronze|Orange|Red|Purple|Blue|Cyan|Green")
		arrIconStarColors = array(strIconStarGold,strIconStarSilver,strIconStarBronze,strIconStarOrange,strIconStarRed,strIconStarPurple,strIconStarBlue,strIconStarCyan,strIconStarGreen)
		strStarColor = split(arrStarColors, "|")

		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""ShowRank""></a>Showing Ranks?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"<ol>" & vbNewLine & _
						"<li>Don't Show Any</li>" & vbNewLine & _
						"<li>Show Rank Only</li>" & vbNewLine & _
						"<li>Show Stars Only</li>" & vbNewLine & _
						"<li>Show Both Stars and Rank</li>" & vbNewLine & _
						"</ol>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""RankColor""></a>Color of Stars?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"You can change the color of stars that show up for each rank of member. (only when the Stars function is turned on)" & vbNewLine & _
						"Available colors for the stars:<br /><br />" & vbNewLine
		for c = 0 to ubound(strStarColor)
			Response.Write	"" & getCurrentIcon(arrIconStarColors(c),"","align=""absmiddle""") & "&nbsp;&nbsp;" & strStarColor(c)
			if c <> ubound(strStarColor) then Response.Write("<br />" & vbNewLine) else Response.Write(vbNewLine)
		next
		Response.Write	"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	case "datetime"
		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""timetype""></a>Time Display?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Choose 24Hr to display all times in military (24 hour) format or 12Hr to display all times in 12 hour format appended with an AM or PM depending on whether it's before or after midday. Default is 24 hour." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""TimeAdjust""></a>Time Adjustment?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Enter either a positive or negative integer value between +12 and 0 and -12. This may come in handy if you are located in one part of the world, and your server is in another, and you need the time displayed in the forum to be converted to a local time for you! (Default value is 0, meaning no adjustment)" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr> " & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""datetype""></a>Date Display?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Choose the format you wish all dates to be displayed in. Default is 12/31/2000 (US Short)." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	case "email"
		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""email""></a>What does E-mail do?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Disabling the E-mail function will turn off any features that involve sending mail. If you don't have an SMTP server of any type, you will want to turn this feature off. If you do have an SMTP (mail) server, however, then also select the type of server you have from the dropdown menu." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""mailserver""></a>What is a Mail Server Address?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"The mail server address is the actual domain name that resolves your mail server. This could be something like:<br />" & vbNewLine & _
						"<b>mail.snitz.com</b><br />" & vbNewLine & _
						"or it could be the same address as the web server:<br />" & vbNewLine & _
						"<b>www.snitz.com</b><br />" & vbNewLine & _
						"Either way, don't put the <b>http://</b> on it.<br />" & vbNewLine & _
						"<br />" & vbNewLine & _
						"<span class=""hlf""><b>NOTE:</b> If you are using CDONTS as a mail server type, you do not need to fill in this field.</span>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""sender""></a>Administrator E-mail Address?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This address is referenced by the forums in a couple ways.<br />" & vbNewLine & _
						"<ol>" & vbNewLine & _
						"<li>When mail is sent, it is sent from this user E-mail Account.</li>" & vbNewLine & _
						"<li>This user is also the point of contact given if there is a problem with these forums.</li>" & vbNewLine & _
						"</ol>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""UniqueEmail""></a>Unique E-mail Address?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Do you want to require each user to have their own E-mail Address?" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""EmailVal""></a>E-mail Validation?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Do you want to require each user to validate their E-mail Address when they first Register and anytime they change their E-mail Address?<br /><br />The user will receive an E-mail with a link in it that will validate that the E-mail Address they entered is a valid E-mail Address." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""EmailFilter""></a>Filter known spam domains?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This allows you to filter out E-mail addresses from a given domain - like all addresses from @example.com.<br /><br />This will prevent people from registering with E-mail addresses at that domain." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""RestrictReg""></a>Restrict Registration?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This allows you to choose who is able to register on your forum by approving or rejecting their registration.<br /><br /><b>Note:</b> You must have the E-mail Validation option turned On to use this feature." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""LogonForMail""></a>Require Logon for sending Mail?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Do you require a user to be logged on before being able to use the <i>Send Topic To a Friend</i> or <i>E-mail Poster</i> options?" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""MaxPostsToEMail""></a>Number of posts to allow sending e-mail?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"To prevent spammers from registering an account and immediately sending your members e-mails through the forum, set this to the number of posts you want someone to have before they can use the forum's e-mail function.<br /><br />Set this to 0 if you want to turn this feature off.<br /><br />" & _
						"There is an Admin overide for this feature. If you want to exempt an individual, edit their profile.<br /><br /><b>Note:</b> this does not affect Admins or Moderators." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""NoMaxPostsToEMail""></a>Error if they don't have enough posts?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This is the message you want someone to see if they don't have enough posts to send an email.<br /><br /><b>Note:</b> ""Number of posts to allow sending e-mail"" must be greater than 0." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
	case "colors"
		Response.Write	"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""fontfacetype""></a>Font Face Type?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Font Face Type changes the way the text in your forum looks. You may want to change this option to match that of the rest of your web site. Some standards are:" & vbNewLine & _
						"<ul>" & vbNewLine & _
						"<li>Arial (nice, clean, legible font)</li>" & vbNewLine & _
						"<li>Courier (a typewriter font)</li>" & vbNewLine & _
						"<li>Helvetica (another clean, legible font)</li>" & vbNewLine & _
						"<li>Sans Serif (Arial & Helvetica are variants of Sans Serif)</li>" & vbNewLine & _
						"<li>Times New Roman (a book-type font)</li>" & vbNewLine & _
						"<li>Verdana (another clean, legible font) (default)</li>" & vbNewLine & _
						"</ul>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""fontsize""></a>What does Font Size mean?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"<ul>" & vbNewLine & _
						"<li>None = Use Browser Default</li>" & vbNewLine & _
						"<li>1 = 8 point font <b>X-Small</b> (default footer size)</li>" & vbNewLine & _
						"<li>2 = 10 point font <b>Small</b> (default font size)</li>" & vbNewLine & _
						"<li>3 = 12 point font <b>Normal</b></li>" & vbNewLine & _
						"<li>4 = 14 point font <b>Large</b> (default header size)</li>" & vbNewLine & _
						"<li>5 = 18 point font <b>X-Large</b></li>" & vbNewLine & _
						"<li>6 = 24 point font <b>XX-Large</b></li>" & vbNewLine & _
						"<li>7 = 36 point font <b>XXX-Large</b></li>" & vbNewLine & _
						"</ul>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""colors""></a>What colors may I use?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"<p>" & vbNewLine & _
						"There are a lot of different colors you can choose from, all of which are listed below:</p>" & vbNewLine & _
						"<blockquote><pre>" & vbNewLine & _
						"<span style=""color:aliceblue"">aliceblue</span>" & vbNewLine & _
						"<span style=""color:antiquewhite"">antiquewhite</span>" & vbNewLine & _
						"<span style=""color:aqua"">aqua</span>" & vbNewLine & _
						"<span style=""color:aquamarine"">aquamarine</span>" & vbNewLine & _
						"<span style=""color:azure"">azure</span>" & vbNewLine & _
						"<span style=""color:beige"">beige</span>" & vbNewLine & _
						"<span style=""color:bisque"">bisque</span>" & vbNewLine & _
						"<span style=""color:black"">black</span>" & vbNewLine & _
						"<span style=""color:blanchedalmond"">blanchedalmond</span>" & vbNewLine & _
						"<span style=""color:blue"">blue</span>" & vbNewLine & _
						"<span style=""color:blueviolet"">blueviolet</span>" & vbNewLine & _
						"<span style=""color:brown"">brown</span>" & vbNewLine & _
						"<span style=""color:burlywood"">burlywood</span>" & vbNewLine & _
						"<span style=""color:cadetblue"">cadetblue</span>" & vbNewLine & _
						"<span style=""color:chartreuse"">chartreuse</span>" & vbNewLine & _
						"<span style=""color:chocolate"">chocolate</span>" & vbNewLine & _
						"<span style=""color:coral"">coral</span>" & vbNewLine & _
						"<span style=""color:cornflowerblue"">cornflowerblue</span>" & vbNewLine & _
						"<span style=""color:cornsilk"">cornsilk</span>" & vbNewLine & _
						"<span style=""color:cyan"">cyan</span>" & vbNewLine & _
						"<span style=""color:darkblue"">darkblue</span>" & vbNewLine & _
						"<span style=""color:darkcyan"">darkcyan</span>" & vbNewLine & _
						"<span style=""color:darkgoldenrod"">darkgoldenrod</span>" & vbNewLine & _
						"<span style=""color:darkgray"">darkgray</span>" & vbNewLine & _
						"<span style=""color:darkgreen"">darkgreen</span>" & vbNewLine & _
						"<span style=""color:darkkhaki"">darkkhaki</span>" & vbNewLine & _
						"<span style=""color:darkmagenta"">darkmagenta</span>" & vbNewLine & _
						"<span style=""color:darkolivegreen"">darkolivegreen</span>" & vbNewLine & _
						"<span style=""color:darkorange"">darkorange</span>" & vbNewLine & _
						"<span style=""color:darkorchid"">darkorchid</span>" & vbNewLine & _
						"<span style=""color:darkred"">darkred</span>" & vbNewLine & _
						"<span style=""color:darksalmon"">darksalmon</span>" & vbNewLine & _
						"<span style=""color:darkseagreen"">darkseagreen</span>" & vbNewLine & _
						"<span style=""color:darkslateblue"">darkslateblue</span>" & vbNewLine & _
						"<span style=""color:darkslategray"">darkslategray</span>" & vbNewLine & _
						"<span style=""color:darkturquoise"">darkturquoise</span>" & vbNewLine & _
						"<span style=""color:darkviolet"">darkviolet</span>" & vbNewLine & _
						"<span style=""color:deeppink"">deeppink</span>" & vbNewLine & _
						"<span style=""color:deepskyblue"">deepskyblue</span>" & vbNewLine & _
						"<span style=""color:dimgray"">dimgray</span>" & vbNewLine & _
						"<span style=""color:dodgerblue"">dodgerblue</span>" & vbNewLine & _
						"<span style=""color:firebrick"">firebrick</span>" & vbNewLine & _
						"<span style=""color:floralwhite"">floralwhite</span>" & vbNewLine & _
						"<span style=""color:forestgreen"">forestgreen</span>" & vbNewLine & _
						"<span style=""color:gainsboro"">gainsboro</span>" & vbNewLine & _
						"<span style=""color:ghostwhite"">ghostwhite</span>" & vbNewLine & _
						"<span style=""color:gold"">gold</span>" & vbNewLine & _
						"<span style=""color:goldenrod"">goldenrod</span>" & vbNewLine & _
						"<span style=""color:gray"">gray</span>" & vbNewLine & _
						"<span style=""color:green"">green</span>" & vbNewLine & _
						"<span style=""color:greenyellow"">greenyellow</span>" & vbNewLine & _
						"<span style=""color:honeydew"">honeydew</span>" & vbNewLine & _
						"<span style=""color:hotpink"">hotpink</span>" & vbNewLine & _
						"<span style=""color:indianred"">indianred</span>" & vbNewLine & _
						"<span style=""color:ivory"">ivory</span>" & vbNewLine & _
						"<span style=""color:khaki"">khaki</span>" & vbNewLine & _
						"<span style=""color:lavender"">lavender</span>" & vbNewLine & _
						"<span style=""color:lavenderblush"">lavenderblush</span>" & vbNewLine & _
						"<span style=""color:lawngreen"">lawngreen</span>" & vbNewLine & _
						"<span style=""color:lemonchiffon"">lemonchiffon</span>" & vbNewLine & _
						"<span style=""color:lightblue"">lightblue</span>" & vbNewLine & _
						"<span style=""color:lightcoral"">lightcoral</span>" & vbNewLine & _
						"<span style=""color:lightcyan"">lightcyan</span>" & vbNewLine & _
						"<span style=""color:lightgoldenrod"">lightgoldenrod</span>" & vbNewLine & _
						"<span style=""color:lightgoldenrodyellow"">lightgoldenrodyellow</span>" & vbNewLine & _
						"<span style=""color:lightgray"">lightgray</span>" & vbNewLine & _
						"<span style=""color:lightgreen"">lightgreen</span>" & vbNewLine & _
						"<span style=""color:lightpink"">lightpink</span>" & vbNewLine & _
						"<span style=""color:lightsalmon"">lightsalmon</span>" & vbNewLine & _
						"<span style=""color:lightseagreen"">lightseagreen</span>" & vbNewLine & _
						"<span style=""color:lightskyblue"">lightskyblue</span>" & vbNewLine & _
						"<span style=""color:lightslateblue"">lightslateblue</span>" & vbNewLine & _
						"<span style=""color:lightslategray"">lightslategray</span>" & vbNewLine & _
						"<span style=""color:lightsteelblue"">lightsteelblue</span>" & vbNewLine & _
						"<span style=""color:lightyellow"">lightyellow</span>" & vbNewLine & _
						"<span style=""color:limegreen"">limegreen</span>" & vbNewLine & _
						"<span style=""color:linen"">linen</span>" & vbNewLine & _
						"<span style=""color:magenta"">magenta</span>" & vbNewLine & _
						"<span style=""color:maroon"">maroon</span>" & vbNewLine & _
						"<span style=""color:mediumaquamarine"">mediumaquamarine</span>" & vbNewLine & _
						"<span style=""color:mediumblue"">mediumblue</span>" & vbNewLine & _
						"<span style=""color:mediumorchid"">mediumorchid</span>" & vbNewLine & _
						"<span style=""color:mediumpurple"">mediumpurple</span>" & vbNewLine & _
						"<span style=""color:mediumseagreen"">mediumseagreen</span>" & vbNewLine & _
						"<span style=""color:mediumslateblue"">mediumslateblue</span>" & vbNewLine & _
						"<span style=""color:mediumspringgreen"">mediumspringgreen</span>" & vbNewLine & _
						"<span style=""color:mediumturquoise"">mediumturquoise</span>" & vbNewLine & _
						"<span style=""color:mediumvioletred"">mediumvioletred</span>" & vbNewLine & _
						"<span style=""color:midnightblue"">midnightblue</span>" & vbNewLine & _
						"<span style=""color:mintcream"">mintcream</span>" & vbNewLine & _
						"<span style=""color:mistyrose"">mistyrose</span>" & vbNewLine & _
						"<span style=""color:moccasin"">moccasin</span>" & vbNewLine & _
						"<span style=""color:navajowhite"">navajowhite</span>" & vbNewLine & _
						"<span style=""color:navy"">navy</span>" & vbNewLine & _
						"<span style=""color:navyblue"">navyblue</span>" & vbNewLine & _
						"<span style=""color:oldlace"">oldlace</span>" & vbNewLine & _
						"<span style=""color:olivedrab"">olivedrab</span>" & vbNewLine & _
						"<span style=""color:orange"">orange</span>" & vbNewLine & _
						"<span style=""color:orangered"">orangered</span>" & vbNewLine & _
						"<span style=""color:orchid"">orchid</span>" & vbNewLine & _
						"<span style=""color:palegoldenrod"">palegoldenrod</span>" & vbNewLine & _
						"<span style=""color:palegreen"">palegreen</span>" & vbNewLine & _
						"<span style=""color:paleturquoise"">paleturquoise</span>" & vbNewLine & _
						"<span style=""color:palevioletred"">palevioletred</span>" & vbNewLine & _
						"<span style=""color:papayawhip"">papayawhi</span>p" & vbNewLine & _
						"<span style=""color:peachpuff"">peachpuff</span>" & vbNewLine & _
						"<span style=""color:peru"">peru</span>" & vbNewLine & _
						"<span style=""color:pink"">pink</span>" & vbNewLine & _
						"<span style=""color:plum"">plum</span>" & vbNewLine & _
						"<span style=""color:powderblue"">powderblue</span>" & vbNewLine & _
						"<span style=""color:purple"">purple</span>" & vbNewLine & _
						"<span style=""color:red"">red</span>" & vbNewLine & _
						"<span style=""color:rosybrown"">rosybrown</span>" & vbNewLine & _
						"<span style=""color:royalblue"">royalblue</span>" & vbNewLine & _
						"<span style=""color:saddlebrown"">saddlebrown</span>" & vbNewLine & _
						"<span style=""color:salmon"">salmon</span>" & vbNewLine & _
						"<span style=""color:sandybrown"">sandybrown</span>" & vbNewLine & _
						"<span style=""color:seagreen"">seagreen</span>" & vbNewLine & _
						"<span style=""color:seashell"">seashell</span>" & vbNewLine & _
						"<span style=""color:sienna"">sienna</span>" & vbNewLine & _
						"<span style=""color:skyblue"">skyblue</span>" & vbNewLine & _
						"<span style=""color:slateblue"">slateblue</span>" & vbNewLine & _
						"<span style=""color:slategray"">slategray</span>" & vbNewLine & _
						"<span style=""color:snow"">snow</span>" & vbNewLine & _
						"<span style=""color:springgreen"">springgreen</span>" & vbNewLine & _
						"<span style=""color:steelblue"">steelblue</span>" & vbNewLine & _
						"<span style=""color:tan"">tan</span>" & vbNewLine & _
						"<span style=""color:thistle"">thistle</span>" & vbNewLine & _
						"<span style=""color:tomato"">tomato</span>" & vbNewLine & _
						"<span style=""color:turquoise"">turquoise</span>" & vbNewLine & _
						"<span style=""color:violet"">violet</span>" & vbNewLine & _
						"<span style=""color:violetred"">violetred</span>" & vbNewLine & _
						"<span style=""color:wheat"">wheat</span>" & vbNewLine & _
						"<span style=""color:white"">white</span>" & vbNewLine & _
						"<span style=""color:whitesmoke"">whitesmoke</span>" & vbNewLine & _
						"<span style=""color:yellow"">yellow</span>" & vbNewLine & _
						"<span style=""color:yellowgreen"">yellowgreen</span>" & vbNewLine & _
						"</pre></blockquote>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""fontdecorations""></a>What are Font Decorations?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"<ul>" & vbNewLine & _
						"<li>none</li>" & vbNewLine & _
						"<li>blink</li>" & vbNewLine & _
						"<li>line-through</li>" & vbNewLine & _
						"<li>overline</li>" & vbNewLine & _
						"<li>underline</li>" & vbNewLine & _
						"</ul>" & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""pagebgimage""></a>What is Page Background Image URL?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"Enter the URL to the location of the background image you would like for your forum." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""columnwidth""></a>How does Column Width Work?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"This sets the width of the column in question. It is not recommended that you change this unless you really know what your doing." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr class=""header"">" & vbNewLine & _
						"<td><a name=""nowrap""></a>What is NOWRAP?</td>" & vbNewLine & _
						"</tr>" & vbNewLine & _
						"<tr>" & vbNewLine & _
						"<td>" & vbNewLine & _
						"NOWRAP prevents the text in a column from auto wrapping. This could be bad if you have people posting long strings of text in the right column (message box), reason being: this would cause an awful long horizontal scroll bar in most cases." & vbNewLine & _
						"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
						"</tr>" & vbNewLine
end select
Response.Write	"</table>" & vbNewLine
Call WriteFooterShort()
%>
