<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/blog.asp" -->
<%
If IsValidString(Request("chosenMonth")) = True AND IsValidString(Request("chosenYear")) = True Then
	' what year and month did the user choose?
	' (again, if the numbers are not valid CInt will give an error)
	theYear = HackerSafe_Filter(CInt( Request("chosenYear")))
	theMonth = HackerSafe_Filter(CInt( Request("chosenMonth")))
Else
	Response.End	
End If

' find first day of the given month...
firstDate = DateSerial( theYear, theMonth, 1 ) ' day 1 of theMonth in theYear
' now comes the "quirk" from the MS documentation:
lastDate = DateSerial( theYear, theMonth + 1, 0 ) 

'SQL = "SELECT * FROM table WHERE theDateField BETWEEN #" & firstDate & "# AND #" & lastDate & "#"
%>
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
if instr(MM_blog_STRING, "Catalog") then
rsArticles.Source = "SELECT BlogID, BlogHeadline, BlogHTML, BlogDate, BlogCat, BlogAuthor, BlogCommentInclude, BlogReadMore, BlogDraft, CatID, CatName, CatDesc, fldAuthorID, fldAuthorRealName, (SELECT COUNT(*) FROM tblComment WHERE tblComment.BlogID = tblBlog.BlogID AND tblComment.CommentInclude = 1) as CommentCount, (SELECT COUNT(*) FROM tblBlog WHERE BlogCat = CatID AND BlogDraft <> 1) as CategoryCount FROM tblBlog, tblCat, tblAuthor WHERE BlogCat = CatID AND Month(BlogDate) = " & theMonth & " AND Year(BlogDate) = " & theYear & " AND tblBlog.BlogAuthor = tblAuthor.fldAuthorID ORDER BY BlogDate DESC"
else
rsArticles.Source = "SELECT BlogID, BlogHeadline, BlogHTML, BlogDate, BlogCat, BlogAuthor, BlogCommentInclude, BlogReadMore, BlogDraft, CatID, CatName, CatDesc, fldAuthorID, fldAuthorRealName, (SELECT COUNT(*) FROM tblComment WHERE tblComment.BlogID = tblBlog.BlogID AND tblComment.CommentInclude = 1) as CommentCount, (SELECT COUNT(*) FROM tblBlog WHERE BlogCat = CatID AND BlogDraft <> 1) as CategoryCount   FROM tblBlog, tblCat, tblAuthor   WHERE BlogCat = CatID AND Month(BlogDate) = " & theMonth & " AND Year(BlogDate) = " & theYear & " AND tblBlog.BlogAuthor = tblAuthor.fldAuthorID AND tblBlog.BlogDraft <> 1 ORDER BY BlogDate DESC"
end if
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
rsBlogSite_numRows = rsBlogSite_numRows + Repeat1__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = -1
Repeat3__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat3__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr">
<head profile="http://gmpg.org/xfn/11">
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | <%=MonthName(theMonth)%><%Response.Write(" " & theYear)%></title>
<meta name="Description" content="<%=(rsBlogSite.Fields.Item("blogDesc").Value)%>" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/andrea/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="wrap" class="group">
<div id="header">
<h1><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a> <%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></h1>
</div>
<div id="menu" class="group">
<div class="corner-TL"></div> <div class="corner-TR"></div>
	<ul>
	<li><a href="default.asp">Home</a></li>
	<li class="page_item page-item-2"><a href="template.asp?pagename=about" title="About">About</a></li>
	<li class="page_item page-item-11"><a href="template_gallery.asp" title="Gallery">Gallery</a></li>
	</ul>
    <div id="feed"><a href="rss.xml">RSS Feed</a></div>
<div class="corner-BL"></div> <div class="corner-BR"></div>
</div>
<div id="content" class="group">
<% 
While ((Repeat3__numRows <> 0) AND (NOT rsArticles.EOF)) 
%>
  <h2 id="post-<%=(rsArticles.Fields.Item("BlogID").Value)%>"><a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"><%=(rsArticles.Fields.Item("BlogHeadline").Value)%></a><a name="<%=(rsArticles.Fields.Item("BlogID").Value)%>" id="<%=(rsArticles.Fields.Item("BlogID").Value)%>"></a></h2>
  <div class="stamp">Posted at <%= DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 3, 1033) %> in <a href="template_archives_cat.asp?cat=<%=(rsArticles.Fields.Item("CatID").Value)%>" title="<%=(rsArticles.Fields.Item("CatDesc").Value)%>"><%=(rsArticles.Fields.Item("CatName").Value)%> (<%=(rsArticles.Fields.Item("CategoryCount").Value)%>)</a></div>
  <div class="main">
  <% if (rsArticles.Fields.Item("BlogReadMore").Value) = 1 Then %>
  <p><%=CropSentence(CI_StripHTML(rsArticles.Fields.Item("BlogHTML").Value), 500, "...")%></p>
  <h4 align="center"><a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>#<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Read More <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>">Read More "<%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"</a></h4>
  <% Else %>
  <%=readmore(rsArticles.Fields.Item("BlogHTML").Value,rsArticles.Fields.Item("BlogID").Value)%>
  <% End If %>
  </div>
  <div class="meta">
  <p>Written by <a href="template_author.asp?id=<%=(rsArticles.Fields.Item("fldAuthorID").Value)%>" title="<%=(rsArticles.Fields.Item("fldAuthorRealName").Value)%>'s Profile"><%=(rsArticles.Fields.Item("fldAuthorRealName").Value)%></a> on <%= DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 1, 1033) %> | <a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>#comments" class="commentsLink">Comments (<%=(rsArticles.Fields.Item("CommentCount").Value)%>)</a></p>
  </div>
  <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  lastdate = DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 1, 1033)
  rsArticles.MoveNext()
Wend
%>
</div>
<div id="sidebar">
<!--#include file="inc_sidebar.asp" -->

</div>

<div id="footer" class="group">
<div class="corner-TL"></div> <div class="corner-TR"></div>
    <div id="top"><a href="#content">Back to top</a></div>
	<p><P>Powered by <a href="http://blog.betaparticle.com" title="Powered by BP Blog 9.0">BP Blog 9.0</a>
		 | <a href="rss.xml">Feed (RSS)</a></P></p>
<div class="corner-BL"></div> <div class="corner-BR"></div>
</div>


</body>
</html>
<%
rsBlogSite.Close()
Set rsBlogSite = Nothing
%>