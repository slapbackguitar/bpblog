<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Buffer = False %>
<%
curpath = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
curpath =  Left(curpath, InstrRev(curpath, "/"))
%>	
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<!--#include file="Connections/blog.asp" -->


<%
Dim rsConfig
Dim rsConfig_numRows
Set rsConfig = Server.CreateObject("ADODB.Recordset")
rsConfig.ActiveConnection = MM_blog_STRING
rsConfig.Source = "SELECT * FROM tblBlogRSS"
rsConfig.CursorType = 0
rsConfig.CursorLocation = 2
rsConfig.LockType = 1
rsConfig.Open()
rsConfig_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Update RSS</title>
	<style type="text/css" media="screen">@import "tabs.css";</style>
</head>
<body>
	<% if Session("MM_Admin") = 1 then %>
	<h3 class="floatright"><a href="?view=1" accesskey="2">User View</a> | <a href="?view=2" accesskey="3">Admin View</a></h3>
	<% end if %>
	<h1><a href="main.asp" accesskey="1">bp blog admin (<%=Session("MM_Username")%>)</a> | <a href="default.asp">Your Blog</a></h1>
	<div id="header">
	<ul id="primary">
		<li><a class="current" href="main.asp">Home (Entries)</a></li>
			<ul id="secondary">
				<li><a href="add_blog.asp">Create a New Entry</a></li>
				<li><a href="rss.asp">Update RSS</a></li>
			</ul>
		<li><a href="user_update.asp?id=<%=Session("MM_UserID")%>">Profile</a></li>
		<li><a href="gallery.asp">Gallery</a></li>
		<% if Session("isAdmin") = 1 then %>
		<li><a href="pages.asp">Pages</a></li>
		<li><a href="cat.asp">Categories</a></li>
		<li><a href="users.asp">Users</a></li>
		<li><a href="layout.asp">Layout</a></li>
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">
<% 
' Timer code
	Dim StopWatch(19) 
	sub StartTimer(x)
		StopWatch(x) = timer
	end sub
	function StopTimer(x)
		EndTime = Timer
		'Watch for the midnight wraparound...
		if EndTime < StopWatch(x) then
			EndTime = EndTime + (86400)
		end if
		StopTimer = EndTime - StopWatch(x)
	end function
	StartTimer 1
%>
   <h2>Update RSS</h2>   
<%
Dim artRec
Dim artRec_numRows
Set artRec = Server.CreateObject("ADODB.Recordset")
artRec.ActiveConnection = MM_blog_STRING
artRec.Source = "SELECT * FROM tblBlog, tblCat, tblAuthor  WHERE BlogCat = CatID  AND tblBlog.BlogAuthor = tblAuthor.fldAuthorID AND tblBlog.BlogDraft <> 1 ORDER BY BlogDate DESC"
artRec.CursorType = 0
artRec.CursorLocation = 2
artRec.LockType = 1
artRec.Open()
artRec_numRows = 0
%>
<%
sFilename = "rss.xml"
%>
<%
	Dim oFSO
	Dim fFile
	' create an instance of the FileSystemObject
	Set oFSO = Server.CreateObject ("Scripting.FileSystemObject")
	' create file
	Set fFile = oFSO.CreateTextFile (Server.MapPath(sFilename))
%>
<%
Dim oRs__currentDate
oRs__currentDate = "Month(Date())"
If (Month(Date()) <> "") Then 
  oRs__currentDate = Month(Date())
End If
%>
<%
Dim rsBlogConfig
Dim rsBlogConfig_numRows
Set rsBlogConfig = Server.CreateObject("ADODB.Recordset")
rsBlogConfig.ActiveConnection = MM_blog_STRING
rsBlogConfig.Source = "SELECT * FROM tblBlogRSS"
rsBlogConfig.CursorType = 0
rsBlogConfig.CursorLocation = 2
rsBlogConfig.LockType = 1
rsBlogConfig.Open()
rsBlogConfig_numRows = 0
%>
<%
function CI_StripHTML(strtext)				
 on error resume next	
	'Strips the HTML tags from strHTML
	
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp
	
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "<(.|\n)+?>"
	
	'Replace all HTML tag matches with the empty string
	strOutput = objRegExp.Replace(strtext, "")

	Set objRegExp = Nothing	
	
	strOutput = replace(strOutput,"&hellip;","...")
	strOutput = replace(strOutput,"&lsquo;","")
	strOutput = replace(strOutput,"&rsquo;","")
	strOutput = replace(strOutput,"&sbquo;","")
	strOutput = replace(strOutput,"&ldquo;", "")
	strOutput = replace(strOutput,"&rdquo;", "")
	strOutput = replace(strOutput,"&bdquo;", "")
	strOutput = replace(strOutput,"&dagger;", "")
	strOutput = replace(strOutput,"&Dagger;", "")
	strOutput = replace(strOutput,"&permil;", "")
	strOutput = replace(strOutput,"&lsaquo;", "")
	strOutput = replace(strOutput,"&rsaquo;", "")
	strOutput = replace(strOutput,"&spades;", "")
	strOutput = replace(strOutput,"&clubs;", "")
	strOutput = replace(strOutput,"&hearts;", "")
	strOutput = replace(strOutput,"&diams;", "")
	strOutput = replace(strOutput,"&oline;", "")
	strOutput = replace(strOutput,"&larr;", "")
	strOutput = replace(strOutput,"&uarr;", "")
	strOutput = replace(strOutput,"&rarr;", "")
	strOutput = replace(strOutput,"&darr;", "")
	strOutput = replace(strOutput,"&trade;", "")
	strOutput = replace(strOutput,"&quot;", "")
	strOutput = replace(strOutput,"&amp;", "")
	strOutput = replace(strOutput,"&frasl;", "")
	strOutput = replace(strOutput,"&lt;", "")
	strOutput = replace(strOutput,"&gt;", "")
	strOutput = replace(strOutput,"&ndash;", "")
	strOutput = replace(strOutput,"&mdash;", "")
	strOutput = replace(strOutput,"&nbsp;", " ")
	strOutput = replace(strOutput,"&iexcl;", "")
	strOutput = replace(strOutput,"&cent;", "")
	strOutput = replace(strOutput,"&pound;", "")
	strOutput = replace(strOutput,"&curren;", "")
	strOutput = replace(strOutput,"&yen;", "")
	strOutput = replace(strOutput,"&brvbar;", "")
	strOutput = replace(strOutput,"&brkbar;", "")
	strOutput = replace(strOutput,"&sect;", "")
	strOutput = replace(strOutput,"&uml;", "")
	strOutput = replace(strOutput,"&die;", "")
	strOutput = replace(strOutput,"&copy;", "")
	strOutput = replace(strOutput,"&ordf;", "")
	strOutput = replace(strOutput,"&laquo;", "")
	strOutput = replace(strOutput,"&not;", "")
	strOutput = replace(strOutput,"&shy;", "")
	strOutput = replace(strOutput,"&reg;", "")
	strOutput = replace(strOutput,"&macr;", "")
	strOutput = replace(strOutput,"&hibar;", "")
	strOutput = replace(strOutput,"&deg;", "")
	strOutput = replace(strOutput,"&plusmn;", "")
	strOutput = replace(strOutput,"&sup2;", "")
	strOutput = replace(strOutput,"&sup3;", "")
	strOutput = replace(strOutput,"&acute;", "")
	strOutput = replace(strOutput,"&micro;", "")
	strOutput = replace(strOutput,"&para;", "")
	strOutput = replace(strOutput,"&middot;", "")
	strOutput = replace(strOutput,"&cedil;", "")
	strOutput = replace(strOutput,"&sup1;", "")
	strOutput = replace(strOutput,"&ordm;", "")
	strOutput = replace(strOutput,"&raquo;", "")
	strOutput = replace(strOutput,"&frac14;", "")
	strOutput = replace(strOutput,"&frac12;", "")
	strOutput = replace(strOutput,"&frac34;", "")
	strOutput = replace(strOutput,"&iquest;", "")
	strOutput = replace(strOutput,"&Agrave;", "")
	strOutput = replace(strOutput,"&Aacute;", "")
	strOutput = replace(strOutput,"&Acirc;", "")
	strOutput = replace(strOutput,"&Atilde;", "")
	strOutput = replace(strOutput,"&Auml;", "")
	strOutput = replace(strOutput,"&Aring;", "")
	strOutput = replace(strOutput,"&AElig;", "")
	strOutput = replace(strOutput,"&Ccedil;", "")
	strOutput = replace(strOutput,"&Egrave;", "")
	strOutput = replace(strOutput,"&Eacute;", "")
	strOutput = replace(strOutput,"&Ecirc;", "")
	strOutput = replace(strOutput,"&Euml;", "")
	strOutput = replace(strOutput,"&Igrave;", "")
	strOutput = replace(strOutput,"&Iacute;", "")
	strOutput = replace(strOutput,"&Icirc;", "")
	strOutput = replace(strOutput,"&Iuml;", "")
	strOutput = replace(strOutput,"&ETH;", "")
	strOutput = replace(strOutput,"&Ntilde;", "")
	strOutput = replace(strOutput,"&Ograve;", "")
	strOutput = replace(strOutput,"&Oacute;", "")
	strOutput = replace(strOutput,"&Ocirc;", "")
	strOutput = replace(strOutput,"&Otilde;", "")
	strOutput = replace(strOutput,"&Ouml;", "")
	strOutput = replace(strOutput,"&times;", "")
	strOutput = replace(strOutput,"&Oslash;", "")
	strOutput = replace(strOutput,"&Ugrave;", "")
	strOutput = replace(strOutput,"&Uacute;", "")
	strOutput = replace(strOutput,"&Ucirc;", "")
	strOutput = replace(strOutput,"&Uuml;", "")
	strOutput = replace(strOutput,"&Yacute;", "")
	strOutput = replace(strOutput,"&THORN;", "")
	strOutput = replace(strOutput,"&szlig;", "")
	strOutput = replace(strOutput,"&agrave;", "")
	strOutput = replace(strOutput,"&aacute;", "")
	strOutput = replace(strOutput,"&acirc;", "")
	strOutput = replace(strOutput,"&atilde;", "")
	strOutput = replace(strOutput,"&auml;", "")
	strOutput = replace(strOutput,"&aring;", "")
	strOutput = replace(strOutput,"&aelig;", "")
	strOutput = replace(strOutput,"&ccedil;", "")
	strOutput = replace(strOutput,"&egrave;", "")
	strOutput = replace(strOutput,"&eacute;", "")
	strOutput = replace(strOutput,"&ecirc;", "")
	strOutput = replace(strOutput,"&euml;", "")
	strOutput = replace(strOutput,"&igrave;", "")
	strOutput = replace(strOutput,"&iacute;", "")
	strOutput = replace(strOutput,"&icirc;", "")
	strOutput = replace(strOutput,"&iuml;", "")
	strOutput = replace(strOutput,"&eth;", "")
	strOutput = replace(strOutput,"&ntilde;", "")
	strOutput = replace(strOutput,"&ograve;", "")
	strOutput = replace(strOutput,"&oacute;", "")
	strOutput = replace(strOutput,"&ocirc;", "")
	strOutput = replace(strOutput,"&otilde;", "")
	strOutput = replace(strOutput,"&ouml;", "")
	strOutput = replace(strOutput,"&divide;", "")
	strOutput = replace(strOutput,"&oslash;", "")
	strOutput = replace(strOutput,"&ugrave;", "")
	strOutput = replace(strOutput,"&uacute;", "")
	strOutput = replace(strOutput,"&ucirc;", "")
	strOutput = replace(strOutput,"&uuml;", "")
	strOutput = replace(strOutput,"&yacute;", "")
	strOutput = replace(strOutput,"&thorn;", "")
	strOutput = replace(strOutput,"&yuml;", "")
	strOutput = replace(strOutput,Chr(10),"")
	strOutput = replace(strOutput,Chr(13),"")
  
  CI_StripHTML = strOutput	
End Function														
%>
<%

%>
<%
Function return_RFC822_Date(myDate, offset)
  Dim myDay, myDays, myMonth, myYear
  Dim myHours, myMonths, mySeconds
	
  myDate = CDate(myDate)
  myDay = WeekdayName(Weekday(myDate),true)
  myDays = Day(myDate)
  myMonth = MonthName(Month(myDate), true)
  myYear = Year(myDate)
  myHours = zeroPad(Hour(myDate), 2)
  myMinutes = zeroPad(Minute(myDate), 2)
  mySeconds = zeroPad(Second(myDate), 2)
	
  return_RFC822_Date = myDay&", "& _
                       myDays&" "& _
                       myMonth&" "& _ 
                       myYear&" "& _
                       myHours&":"& _
                       myMinutes&":"& _
                       mySeconds&" "& _ 
                       offset
End Function 
Function zeroPad(m, t)
  zeroPad = String(t-Len(m),"0")&m
End Function
%>
<%
	sSiteTitle = (rsBlogConfig.Fields.Item("blogTitle").Value)
	sSiteDescr = (rsBlogConfig.Fields.Item("blogDesc").Value)
	sSiteURL = (rsBlogConfig.Fields.Item("blogURL").Value)
	if right(sSiteURL, 1) <> "/" then
		sSiteURL = sSiteURL & "/"
	end if
	sSiteDetails = ""
	sImageURL = (rsBlogConfig.Fields.Item("blogImage").Value)
	sFurtherReading = ""
	sAuthorNames = (rsBlogConfig.Fields.Item("blogAuthor").Value)
	sAuthorEmails = (rsBlogConfig.Fields.Item("blogEmail").Value)
%>
<%
fFile.WriteLine ("<?xml version=""1.0"" encoding=""utf-8""?>")
'fFile.WriteLine ("<?xml-stylesheet type=""text/css"" href=""" & sSiteURL &  "rss.css"" ?>")
fFile.WriteLine ("<!--  RSS generated by " & sSiteTitle & " on " & Now() & " -->")
fFile.WriteLine ("<rss version=""2.0"">")
fFile.WriteLine ("   <channel>")
fFile.WriteLine("	<title>" & sSiteTitle & "</title>")
fFile.WriteLine("	<link>" & sSiteURL & "</link>")
fFile.WriteLine("	<description>" & sSiteDescr & "</description>")
fFile.WriteLine("	<language>en-US</language>")
fFile.WriteLine("	<lastBuildDate>" & return_RFC822_Date(Now(), "GMT") & "</lastBuildDate>")
fFile.WriteLine("	<category domain=""" & sSiteURL & """>" & sSiteTitle & "</category>")
fFile.WriteLine("	<generator>BP Blog 8.0</generator>")
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 10
Repeat1__index = 0
artRec_numRows = artRec_numRows + Repeat1__numRows
%>
<% 
While ((Repeat1__numRows <> 0) AND (NOT artRec.EOF)) 
%>
<%
	fFile.WriteLine ("	<item>")
	fFile.WriteLine ("		<title>" & artRec("BlogHeadline") & "</title>")
	fFile.WriteLine ("		<link>" & sSiteURL & "template_permalink.asp?id=" & artRec.Fields.Item("BlogID").Value & "</link>")
	fFile.WriteLine ("		<guid isPermaLink=""true"">" & sSiteURL & "template_permalink.asp?id=" & artRec.Fields.Item("BlogID").Value & "</guid>")
	fFile.WriteLine ("		<description>" & (CropSentence(CI_StripHTML(artRec.Fields.Item("BlogHTML").Value), 250, "...")) & "</description>")
	fFile.WriteLine ("		<pubDate>" & return_RFC822_Date(artRec.Fields.Item("BlogDate").Value, "GMT") & "</pubDate>")
	fFile.WriteLine("		<category domain=""" & sSiteURL & "template_archives_cat.asp?cat=" & artRec.Fields.Item("CatID").Value & """>" & artRec.Fields.Item("CatName").Value & "</category>")
	fFile.WriteLine ("	</item>")
%>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  artRec.MoveNext()
Wend
%>
<%
' Write out the End Channel Data tag and End RSS tag.
fFile.WriteLine ("	</channel>")
fFile.WriteLine ("</rss>")
	fFile.Close 
	Set fFile = Nothing
	Set oFSO = Nothing
%>
<%
rsBlogConfig.Close()
Set rsBlogConfig = Nothing
%>
<%
artRec.Close()
Set artRec = Nothing
%>
<p>RSS published in <%=abs(round(StopTimer(1), 2))%> seconds. </p>
<p>Back to the <a href="main.asp">main admin page</a>. </p>
		</div>
	</div>
</body>
</html>
<%
rsConfig.Close()
Set rsConfig = Nothing
%>

