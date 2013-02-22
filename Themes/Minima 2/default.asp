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
rsBlogSite_numRows = rsBlogSite_numRows + Repeat1__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = rsBlogSite.Fields.Item("BlogPosts").Value
Repeat3__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat3__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | <%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></title>
<meta name="Description" content="<%=(rsBlogSite.Fields.Item("blogDesc").Value)%>" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/Minima 2/styles-site.css" rel="stylesheet" type="text/css" />
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
While ((Repeat3__numRows <> 0) AND (NOT rsArticles.EOF)) 
%>
<% if lastdate <> DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 1, 1033) then%>
<h2 class="date-header"><%= DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 1, 1033) %></h2>
<%end if%>
<!-- Begin post -->
<div class="post">
  <h3 class="post-title" id="<%=(rsArticles.Fields.Item("BlogID").Value)%>"><a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"><%=(rsArticles.Fields.Item("BlogHeadline").Value)%></a><a name="<%=(rsArticles.Fields.Item("BlogID").Value)%>" id="<%=(rsArticles.Fields.Item("BlogID").Value)%>"></a></h3>
  <div class="post-body">
  <% if (rsArticles.Fields.Item("BlogReadMore").Value) = 1 Then %>
  <p><%=CropSentence(CI_StripHTML(rsArticles.Fields.Item("BlogHTML").Value), 500, "...")%></p>
  <h4 align="center"><a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>#<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Read More <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>">Read More "<%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"</a></h4>
  <% Else %>
  <%=readmore(rsArticles.Fields.Item("BlogHTML").Value,rsArticles.Fields.Item("BlogID").Value)%>
  <% End If %>
  </div>
  <p class="post-footer">Posted by <a href="template_author.asp?id=<%=(rsArticles.Fields.Item("fldAuthorID").Value)%>" title="<%=(rsArticles.Fields.Item("fldAuthorRealName").Value)%>'s Profile"><%=(rsArticles.Fields.Item("fldAuthorRealName").Value)%></a> at <a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"><%= DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 3, 1033) %></a> in <a href="template_archives_cat.asp?cat=<%=(rsArticles.Fields.Item("CatID").Value)%>" title="<%=(rsArticles.Fields.Item("CatDesc").Value)%>"><%=(rsArticles.Fields.Item("CatName").Value)%> (<%=(rsArticles.Fields.Item("CategoryCount").Value)%>)</a> | <a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>#comments">Comments (<%=(rsArticles.Fields.Item("CommentCount").Value)%>)</a>
  </p>
</div>
<!-- End post -->
  <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  lastdate = DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 1, 1033)
  rsArticles.MoveNext()
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