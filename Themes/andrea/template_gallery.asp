<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/blog.asp" -->
<%
Dim rsBlogSite
Dim rsBlogSite_numRows

Set rsBlogSite = Server.CreateObject("ADODB.Recordset")
rsBlogSite.ActiveConnection = MM_blog_STRING
rsBlogSite.Source = "SELECT blogURL, blogTitle, blogSubTitle, blogDesc, blogPosts, blogLayout FROM tblBlogRSS"
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
Dim rsGalleryConfig
Dim rsGalleryConfig_numRows

Set rsGalleryConfig = Server.CreateObject("ADODB.Recordset")
rsGalleryConfig.ActiveConnection = MM_blog_STRING
rsGalleryConfig.Source = "SELECT fldGalleryTitleThumb, fldGalleryThumb FROM tblGalleryConfig"
rsGalleryConfig.CursorType = 0
rsGalleryConfig.CursorLocation = 2
rsGalleryConfig.LockType = 1
rsGalleryConfig.Open()

rsGalleryConfig_numRows = 0
%>
<%
Dim rsGallery
Dim rsGallery_numRows

Set rsGallery = Server.CreateObject("ADODB.Recordset")
rsGallery.ActiveConnection = MM_blog_STRING
rsGallery.Source = "SELECT fldGalleryID, fldGalleryTitle, fldGalleryDesc, fldGalleryPic, fldGalleryCreated, fldGalleryUser, fldAuthorRealName, fldAuthorID FROM tblGallery, tblAuthor WHERE fldGalleryUser = fldAuthorID ORDER BY fldGalleryCreated DESC"
rsGallery.CursorType = 0
rsGallery.CursorLocation = 2
rsGallery.LockType = 1
rsGallery.Open()

rsGallery_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsBlogSite_numRows = rsBlogSite_numRows + Repeat1__numRows
%>
<%
Dim Repeat5__numRows
Dim Repeat5__index

Repeat5__numRows = -1
Repeat5__index = 0
rsGallery_numRows = rsGallery_numRows + Repeat5__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 20
Repeat3__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat3__numRows
%>
<%
galleryroot = Replace(LCase(Request.ServerVariables("PATH_INFO")), "template_gallery.asp", "") & "images/"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr">
<head profile="http://gmpg.org/xfn/11">
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | Gallery</title>
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
<div class="main">  
<% 
While ((Repeat5__numRows <> 0) AND (NOT rsGallery.EOF)) 
%>
<div class="gallery">
<div class="thumb"><a href="template_gallery_detail.asp?fldGalleryID=<%=(rsGallery.Fields.Item("fldGalleryID").Value)%>" title="View this gallery"><img src="thumbnailimage.aspx?filename=<%=galleryroot%><%=(rsGallery.Fields.Item("fldGalleryID").Value)%>/<%=(rsGallery.Fields.Item("fldGalleryPic").Value)%>&width=<%=(rsGalleryConfig.Fields.Item("fldGalleryTitleThumb").Value)%>" border="0" /></a></div>
<h3><a href="template_gallery_detail.asp?fldGalleryID=<%=(rsGallery.Fields.Item("fldGalleryID").Value)%>" title="View this gallery"><%=(rsGallery.Fields.Item("fldGalleryTitle").Value)%>
</a></h3>
<p><%=(rsGallery.Fields.Item("fldGalleryDesc").Value)%></p>
<p>Added on <%= DoDateTime((rsGallery.Fields.Item("fldGalleryCreated").Value), 2, -1) %> by <a href="template_author.asp?id=<%=(rsGallery.Fields.Item("fldAuthorID").Value)%>" title="<%=(rsGallery.Fields.Item("fldAuthorRealName").Value)%>'s Profile"><%=(rsGallery.Fields.Item("fldAuthorRealName").Value)%></a></p>
</div>
<% 
  Repeat5__index=Repeat5__index+1
  Repeat5__numRows=Repeat5__numRows-1
  rsGallery.MoveNext()
Wend
%>
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