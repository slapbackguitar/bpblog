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
If (Request.QueryString("PageName") <> "") AND (Len(Request.QueryString("PageName")) < 10) AND IsValidString(Request.QueryString("PageName")) = True  Then 
  rsPage__MMColParam = HackerSafe_Filter(Request.QueryString("PageName"))
Else
	Response.End	  
End If
%>
<%
Dim rsPage
Dim rsPage_numRows

Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.ActiveConnection = MM_blog_STRING
rsPage.Source = "SELECT PageName, PageTitle, PageHTML, PageDate FROM tblPage WHERE PageName = '" + Replace(rsPage__MMColParam, "'", "''") + "'"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | <%=(rsPage.Fields.Item("PageTitle").Value)%></title>
<meta name="Description" content="<%=(CropSentence(CI_StripHTML(rsPage.Fields.Item("PageTitle").Value), 250, "..."))%>" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/Blue Minima/styles-site.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<DIV id=container>
<DIV id=header>
<h1 id="blog-title"><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a></h1>
<p id="description"><%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></p>
</DIV>
<DIV id=content>
<DIV id=main>
<!-- Begin post -->
<div class="post">
  <h3 class="post-title"><%=(rsPage.Fields.Item("PageTitle").Value)%><% if Session("MM_Username") <> "" AND Session("MM_UserID") = rsArticles.Fields.Item("fldAuthorID").Value OR Session("MM_isAdmin") = 1 Then %> | <a href="update_page.asp?pagename=<%=(rsPage.Fields.Item("pagename").Value)%>">Edit this post</a><% end if %></h3>
  <div class="post-body">  
<%=(rsPage.Fields.Item("PageHTML").Value)%>
</div>
</div><!-- End post -->
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
rsPage.Close()
Set rsPage = Nothing
%>