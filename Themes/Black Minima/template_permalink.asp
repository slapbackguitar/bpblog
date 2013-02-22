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
Dim rs_cat__MMColParam
If Len(Request("id")) < 5 AND IsValidString(Request("id")) = True Then
         rs_cat__MMColParam = HackerSafe_Filter(Request("id"))
Else
	Response.End       
End if

%>
<%
Dim rsArticles
Dim rsArticles_numRows

Set rsArticles = Server.CreateObject("ADODB.Recordset")
rsArticles.ActiveConnection = MM_blog_STRING
rsArticles.Source = "SELECT BlogID, BlogHeadline, BlogHTML, BlogDate, BlogCat, BlogAuthor, BlogCommentInclude, BlogReadMore, BlogDraft, CatID, CatName, CatDesc, fldAuthorID, fldAuthorRealName, (SELECT COUNT(*) FROM tblComment WHERE tblComment.BlogID = tblBlog.BlogID AND tblComment.CommentInclude = 1) as CommentCount, (SELECT COUNT(*) FROM tblBlog WHERE BlogCat = CatID) as CategoryCount FROM tblBlog, tblCat, tblAuthor WHERE BlogID = " & rs_cat__MMColParam & " AND BlogCat = CatID  AND tblBlog.BlogAuthor = tblAuthor.fldAuthorID ORDER BY BlogDate DESC"
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

Repeat3__numRows = 20
Repeat3__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat3__numRows
%>
<%
If cstr(Request.Form("txtName"))<>"" Then
If Request.form("remember") ="1" Then
Response.Cookies("ckName") = Request.Form("txtName")
Response.Cookies("ckURL") = Request.Form("txtURL")
Response.Cookies("ckEmail") = Request.Form("txtEmail")
Response.Cookies("ckRemember") = "1"
Response.Cookies("ckName").Expires = Date + 30
Response.Cookies("ckURL").Expires = Date + 30
Response.Cookies("ckEmail").Expires = Date + 30
Response.Cookies("ckRemember").expires = Date + 30
Else
Response.Cookies("ckName") = ""
Response.Cookies("ckURL") = ""
Response.Cookies("ckEmail") = ""
Response.Cookies("ckRemember") = ""
End If
End If
%>
<%
Dim rsComments__MMColParam
rsComments__MMColParam = "0"
%>
<%
Dim rsComments__MMColParam2
rsComments__MMColParam2 = "0"
If (Request.QueryString("id")  <> "") AND Len(Request.QueryString("id")) < 5 Then 
  rsComments__MMColParam2 = rs_cat__MMColParam
End If
%>
<%
Dim rsComments
Dim rsComments_numRows

Set rsComments = Server.CreateObject("ADODB.Recordset")
rsComments.ActiveConnection = MM_blog_STRING
rsComments.Source = "SELECT commentID, blogID, commentDate, commentName, commentEmail, commentURL, commentHTML, commentInclude FROM tblComment  WHERE commentInclude <> " + Replace(rsComments__MMColParam, "'", "''") + " AND blogID = " + Replace(rsComments__MMColParam2, "'", "''") + "  ORDER BY commentID ASC"
rsComments.CursorType = 0
rsComments.CursorLocation = 2
rsComments.LockType = 1
rsComments.Open()

rsComments_numRows = 0
%>
<%
Dim Repeat5__numRows
Dim Repeat5__index

Repeat5__numRows = -1
Repeat5__index = 0
rsComments_numRows = rsComments_numRows + Repeat5__numRows
%>
<%
theBlogHTML = rsArticles.Fields.Item("BlogHTML").Value
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=(rsArticles.Fields.Item("BlogHeadline").Value)%></title>
<meta name="Description" content="<%=(CropSentence(CI_StripHTML(theBlogHTML), 250, "..."))%>" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script language="JavaScript" type="text/javascript" src="scripts.js"></script>
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/Black Minima/styles-site.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<DIV id=container>
<DIV id=header>
<h1 id="blog-title"><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a></h1>
<p id="description"><%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></p>
</DIV>
<DIV id=content>
<DIV id=main>
<h2 class="date-header"><%= DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 1, 1033) %></h2>
<div class="post"><h3 class="post-title" id="<%=(rsArticles.Fields.Item("BlogID").Value)%>"><a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"><%=(rsArticles.Fields.Item("BlogHeadline").Value)%></a><% if Session("MM_Username") <> "" AND Session("MM_UserID") = rsArticles.Fields.Item("fldAuthorID").Value OR Session("MM_isAdmin") = 1 Then %> | <a href="update_blog.asp?passID=<%=(rsArticles.Fields.Item("BlogID").Value)%>">Edit this post</a><% end if %></h3><a name="<%=(rsArticles.Fields.Item("BlogID").Value)%>" id="<%=(rsArticles.Fields.Item("BlogID").Value)%>"></a>
  <div class="post-body">
  <%=readmore(theBlogHTML,0)%>
  </div>
<p class="post-footer">Posted by <a href="template_author.asp?id=<%=(rsArticles.Fields.Item("fldAuthorID").Value)%>" title="<%=(rsArticles.Fields.Item("fldAuthorRealName").Value)%>'s Profile"><%=(rsArticles.Fields.Item("fldAuthorRealName").Value)%></a> at <a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsArticles.Fields.Item("BlogHeadline").Value)%>"><%= DoDateTime((rsArticles.Fields.Item("BlogDate").Value), 3, 1033) %></a> in <a href="template_archives_cat.asp?cat=<%=(rsArticles.Fields.Item("CatID").Value)%>" title="<%=(rsArticles.Fields.Item("CatDesc").Value)%>"><%=(rsArticles.Fields.Item("CatName").Value)%> (<%=(rsArticles.Fields.Item("CategoryCount").Value)%>)</a> | <a href="template_permalink.asp?id=<%=(rsArticles.Fields.Item("BlogID").Value)%>#comments">Comments (<%=(rsArticles.Fields.Item("CommentCount").Value)%>)</a></p>
<h2>Comments</h2><a name="comments" id="comments"></a>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsComments.EOF)) 
CommentCount = CommentCount + 1
%>
<% If CommentCount MOD 2 = 0 Then %><div class="commentalt"><% end if %>
<% If  Len(rsComments.Fields.Item("commentURL").Value) > 12 Then %>
<h4><%=CommentCount%>. <a href="<%=(rsComments.Fields.Item("commentURL").Value)%>" rel="nofollow" title="Visit <%=(rsComments.Fields.Item("commentName").Value)%>'s Website" target="_blank"><%=(rsComments.Fields.Item("commentName").Value)%></a> said...<a name="#<%=(rsComments.Fields.Item("commentID").Value)%>" id="<%=(rsComments.Fields.Item("commentID").Value)%>"></a></h4>
<% Else %>
<h4><%=CommentCount%>. <%=(rsComments.Fields.Item("commentName").Value)%>  said...<a name="#<%=(rsComments.Fields.Item("commentID").Value)%>" id="<%=(rsComments.Fields.Item("commentID").Value)%>"></a></h4>
<% End If %>
<p><%=MakeHyperlink(rsComments.Fields.Item("commentHTML").Value)%></p>
<p class="post-footer"><a href="#<%=(rsComments.Fields.Item("commentID").Value)%>" title="Comment Permalink"><%=(rsComments.Fields.Item("commentDate").Value)%></a></p>
<% If CommentCount MOD 2 = 0 Then %></div><% end if %>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsComments.MoveNext()
Wend
%>
        <% if (rsArticles.Fields.Item("BlogCommentInclude").Value) = 1 then %>
<form action="comments.asp" method="post" name="form1" id="form1">

<table width="90%"  border="0" cellspacing="2" cellpadding="3">
<tr>
<td align="right" valign="top">Name</td>
<td align="left" valign="middle">
<input name="txtName" type="text" id="txtName" value="" />
</td>
</tr>
<tr>
<td align="right" valign="top">URL</td>
<td align="left" valign="middle">
<input name="txtURL" type="text" id="txtURL" value="http://" /></td>
</tr>
<tr>
<td align="right" valign="top">Email</td>
<td align="left" valign="middle">
<input name="txtEmail" type="text" id="txtEmail" value="" />
<br />
Email address is not published</td>
</tr>
<tr>
<td align="right" valign="top">Remember Me</td>
<td align="left" valign="middle"><input name="remember" type="checkbox" id="remember" value="1" checked="checked" /></td>
</tr>
<tr>
<td align="right" valign="top">Comments</td>
<td align="left" valign="middle">
<textarea name="textarea" rows="5"></textarea></td>
</tr>
<tr>
<td valign="top">
<input name="hiddenField" type="hidden" value="<%= Request.Querystring("id") %>" /></td>
<td align="left" valign="middle">
<script type="text/javascript">
function reloadCAPTCHA() {
document.getElementById('CAPTCHA').src='aspcaptcha.asp?'+Date();
}
</script>

<p><img id="CAPTCHA" src='aspcaptcha.asp' 
alt='CAPTCHA' width='86' height='21' /> 
<a href="javascript:reloadCAPTCHA();">Reload</a>
<br />Write the characters in the image above<br />
<input name='strCAPTCHA' type='text' 
id='strCAPTCHA' maxlength='8' /></p>
<input name="Submit" type="submit" onClick="YY_checkform('form1','txtName','#q','0','Name is required','textarea','1','1','Comment Required');return document.MM_returnValue" value="Comment" /></td>
</tr>
</table>
<input type="hidden" name="MM_insert" value="form1" />
<script language="JavaScript" type="text/javascript">
<!-- Hide from older browsers

  var txtName = getCookie ("txtName");
  if (txtName == null) txtName = "";
  document.form1.txtName.value = txtName;
    var txtURL = getCookie ("txtURL");
  if (txtURL == null) txtURL = "";
  document.form1.txtURL.value = txtURL;
    var txtEmail = getCookie ("txtEmail");
  if (txtEmail == null) txtEmail = "";
  document.form1.txtEmail.value = txtEmail;

// Stop hiding -->
</script>
</form>
<% else %>
<p>Commenting has been turned off for this entry.</p>
<% end if %>
</div>
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