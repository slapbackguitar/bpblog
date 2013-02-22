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
Dim rsPage__MMColParam
rsPage__MMColParam = "1"
If (Request.QueryString("id") <> "") AND (Len(Request.QueryString("id")) < 5) AND IsValidString(Request.QueryString("id")) = True  Then 
  rsPage__MMColParam = HackerSafe_Filter(Request.QueryString("id"))
Else
	Response.End	  
End If
%>
<%
Dim rsPage
Dim rsPage_numRows

Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.ActiveConnection = MM_blog_STRING
rsPage.Source = "SELECT fldAuthorID, fldAuthorRealName, fldAuthorEmail, fldAuthorWebsite, fldAuthorBlurb FROM tblAuthor WHERE fldAuthorID = " & rsPage__MMColParam
rsPage.CursorType = 0
rsPage.CursorLocation = 2
rsPage.LockType = 1
rsPage.Open()

rsPage_numRows = 0
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

Repeat3__numRows = 20
Repeat3__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat3__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr">
<head profile="http://gmpg.org/xfn/11">
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | <%=(rsPage.Fields.Item("fldAuthorRealName").Value)%>'s Profile</title>
<meta name="Description" content="<%=(CropSentence(CI_StripHTML(rsPage.Fields.Item("fldAuthorBlurb").Value), 250, "..."))%>" />
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
  <h2><%=(rsPage.Fields.Item("fldAuthorRealName").Value)%>'s Profile<% if Session("MM_Username") <> "" AND Session("MM_UserID") = rsPage.Fields.Item("fldAuthorID").Value Then %> | <a href="user_update.asp?id=<%=(rsPage.Fields.Item("fldAuthorID").Value)%>">Edit this profile</a><% end if %></h2>
  <h4><strong>Website: </strong><a href="<%=(rsPage.Fields.Item("fldAuthorWebsite").Value)%>" target="_blank"><%=(rsPage.Fields.Item("fldAuthorWebsite").Value)%></a></h4>
<div class="main">  
<%=(rsPage.Fields.Item("fldAuthorBlurb").Value)%>
</div>
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