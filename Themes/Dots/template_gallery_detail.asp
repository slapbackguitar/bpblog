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
Dim rsGalleryDetail__MMColParam
rsGalleryDetail__MMColParam = "1"
If (Request.QueryString("fldGalleryID") <> "") AND (Len(Request.QueryString("fldGalleryID")) < 5) AND (IsValidString(Request.QueryString("fldGalleryID")) = True) Then 
  rsGalleryDetail__MMColParam = HackerSafe_Filter(Request.QueryString("fldGalleryID"))
Else
	Response.End  
End If
%>
<%
Dim rsGalleryDetail
Dim rsGalleryDetail_numRows

Set rsGalleryDetail = Server.CreateObject("ADODB.Recordset")
rsGalleryDetail.ActiveConnection = MM_blog_STRING
rsGalleryDetail.Source = "SELECT fldGalleryID, fldGalleryTitle, fldGalleryDesc, fldGalleryPic, fldGalleryCreated, fldGalleryUser, fldAuthorRealName, fldAuthorID FROM tblGallery, tblAuthor WHERE fldGalleryUser = fldAuthorID AND fldGalleryID = " + Replace(rsGalleryDetail__MMColParam, "'", "''") + ""
rsGalleryDetail.CursorType = 0
rsGalleryDetail.CursorLocation = 2
rsGalleryDetail.LockType = 1
rsGalleryDetail.Open()

rsGalleryDetail_numRows = 0
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
<%
galleryroot = Replace(LCase(Request.ServerVariables("PATH_INFO")), "template_gallery_detail.asp", "") & "images/"
%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | Gallery | <%=(rsGalleryDetail.Fields.Item("fldGalleryTitle").Value)%></title>
<meta name="Description" content="<%=(rsGalleryDetail.Fields.Item("fldGalleryDesc").Value)%>" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/Dots/styles-site.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<DIV id=container>
<DIV id=header>
<h1 id="blog-title"><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a></h1>
<p id="description"><%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></p>
</DIV>
<DIV id=content>
<DIV id=main>
<div class="gallery">
<div class="thumb"><a title="<%=(rsGalleryDetail.Fields.Item("fldGalleryTitle").Value)%>" rel="lightbox[<%=(rsGalleryDetail.Fields.Item("fldGalleryID").Value)%>]" href="<%=galleryroot%><%=(rsGalleryDetail.Fields.Item("fldGalleryID").Value)%>/<%=(rsGalleryDetail.Fields.Item("fldGalleryPic").Value)%>"><img src="thumbnailimage.aspx?filename=<%=galleryroot%><%=(rsGalleryDetail.Fields.Item("fldGalleryID").Value)%>/<%=(rsGalleryDetail.Fields.Item("fldGalleryPic").Value)%>&amp;width=<%=(rsGalleryConfig.Fields.Item("fldGalleryTitleThumb").Value)%>" alt="View Full Size" border="0" /></a></div>
<h3><%=(rsGalleryDetail.Fields.Item("fldGalleryTitle").Value)%></h3>
<p><%=(rsGalleryDetail.Fields.Item("fldGalleryDesc").Value)%></p>
<p class="post-footer">Added on <%= DoDateTime((rsGalleryDetail.Fields.Item("fldGalleryCreated").Value), 2, -1) %> by <a href="template_author.asp?id=<%=(rsGalleryDetail.Fields.Item("fldAuthorID").Value)%>" title="<%=(rsGalleryDetail.Fields.Item("fldAuthorRealName").Value)%>'s Profile"><%=(rsGalleryDetail.Fields.Item("fldAuthorRealName").Value)%></a></p>
</div>
<div class="gallery">
<% 
thepath = galleryroot & (rsGalleryDetail.Fields.Item("fldGalleryID").Value) & "/"
Recurse(Server.MapPath(thepath))
%>
<% 
Sub Recurse(Path)
	
	Dim fso, Root, WindowsFolder, Files, _
		Folders, File, i, FoldersArray(100)
		
	Const ForReading = 1
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set Root = fso.getfolder(Path)
	Set Files = Root.Files
	Set Folders = Root.SubFolders
	For Each File In Files 
	if file.name <> (rsGalleryDetail.Fields.Item("fldGalleryPic").Value) then
	%>
<div class="thumb"><a href="<%=thepath & file.name%>" title="<%=(rsGalleryDetail.Fields.Item("fldGalleryTitle").Value)%>" rel="lightbox[<%=(rsGalleryDetail.Fields.Item("fldGalleryID").Value)%>]"><img src="thumbnailimage.aspx?filename=<%=thepath & file.name%>&amp;width=<%=(rsGalleryConfig.Fields.Item("fldGalleryThumb").Value)%>" alt="View Full Size" /></a></div>
	<% end if
	Next
		 	
	For i = 0 To UBound(FoldersArray)
		if FoldersArray(i) <> "" Then 
			Recurse FoldersArray(i)				
		Else
			Exit For
		End if
	Next
	
End Sub
%>
</div>
<br clear="all" />
<div align="center"><a href="template_gallery.asp">Back to Gallery List </a></div>
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
rsGalleryConfig.Close()
Set rsGalleryConfig = Nothing
%>
<%
rsGalleryDetail.Close()
Set rsGalleryDetail = Nothing
%>