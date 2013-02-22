<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/blog.asp" -->
<%
Dim rsBlogSite
Dim rsBlogSite_numRows

Set rsBlogSite = Server.CreateObject("ADODB.Recordset")
rsBlogSite.ActiveConnection = MM_blog_STRING
rsBlogSite.Source = "SELECT blogTitle, blogSubTitle, blogDesc, blogPosts, blogLayout FROM tblBlogRSS"
rsBlogSite.CursorType = 0
rsBlogSite.CursorLocation = 2
rsBlogSite.LockType = 1
rsBlogSite.Open()

rsBlogSite_numRows = 0
%>
<%
Dim rsArticles
Dim rsArticles_numRows

Set rsArticles = Server.CreateObject("ADODB.Recordset")
rsArticles.ActiveConnection = MM_blog_STRING
rsArticles.Source = "SELECT BlogID, BlogHeadline, BlogHTML, BlogDate, BlogCat, BlogAuthor, BlogCommentInclude, BlogReadMore, BlogDraft, CatID, CatName, CatDesc, fldAuthorID, fldAuthorRealName, (SELECT COUNT(*) FROM tblComment WHERE tblComment.BlogID = tblBlog.BlogID AND tblComment.CommentInclude = 1) as CommentCount, (SELECT COUNT(*) FROM tblBlog WHERE BlogCat = CatID AND BlogDraft <> 1) as CategoryCount  FROM tblBlog, tblCat, tblAuthor  WHERE BlogCat = CatID  AND tblBlog.BlogAuthor = tblAuthor.fldAuthorID AND tblBlog.BlogDraft <> 1 ORDER BY BlogDate DESC"
rsArticles.CursorType = 0
rsArticles.CursorLocation = 2
rsArticles.LockType = 1
rsArticles.Open()

rsArticles_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | Archives</title>
<meta name="Description" content="Archives by Month" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/White Minima/styles-site.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<DIV id=container>
<DIV id=header>
<h1 id="blog-title"><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a></h1>
<p id="description"><%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></p>
</DIV>
<DIV id=content>
<DIV id=main>
<% 
function DayIt(theday)
	if theday < 10 then
		theday = "0" & CStr(theday)
	end if
	DayIt = theDay
end function
While ((Repeat1__numRows <> 0) AND (NOT rsArticles.EOF)) 
MonthHeader = MonthName(Month(rsArticles.Fields.Item("BlogDate").Value)) & " " & Year(rsArticles.Fields.Item("BlogDate").Value)
%>
<% if CurrentMonthHeader <> MonthHeader Then %>
<h3 class="post-title"><a href="template_archives.asp?chosenYear=<%=Year(rsArticles.Fields.Item("BlogDate").Value)%>&chosenMonth=<%=Month(rsArticles.Fields.Item("BlogDate").Value)%>"><%=MonthHeader%></a></h3>
<ul><% end if %>
<li class="archives"><%=DayIt(Day(rsArticles.Fields.Item("BlogDate").Value))%>: <a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"><%=(rsArticles.Fields.Item("BlogHeadline").Value)%></a> (<%=(rsArticles.Fields.Item("CommentCount").Value)%>)</li>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsArticles.MoveNext()
  CurrentMonthHeader = MonthHeader
  if NOT rsArticles.EOF then
  if CurrentMonthHeader <> (MonthName(Month(rsArticles.Fields.Item("BlogDate").Value)) & " " & Year(rsArticles.Fields.Item("BlogDate").Value)) Then Response.Write("</ul>") end if
  end if
Wend
%>
</DIV>
<DIV id=sidebar>
<!--#include file="../../inc_sidebar.asp" -->
</DIV>
 </DIV>
<DIV id=footer><P>Powered by <a href="http://blog.betaparticle.com" title="Powered by BP Blog 9.0">BP Blog 9.0</a>
		 | <a href="rss.xml">Feed (RSS)</a></P>
</DIV>	
</DIV>	 
</BODY>
</HTML>
<%
rsBlogSite.Close()
Set rsBlogSite = Nothing
%>
<%
rsArticles.Close()
Set rsArticles = Nothing
%>