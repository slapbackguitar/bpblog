<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
Dim rsComments_Pending

Set rsComments_Pending = Server.CreateObject("ADODB.Recordset")
rsComments_Pending.ActiveConnection = MM_blog_STRING
if Session("isAdmin") = 1 then
	rsComments_Pending.Source = "SELECT Count(*) as CommentsPendingCount FROM tblComment WHERE commentInclude = 0"
elseif Session("isAdmin") = 0 then
	rsComments_Pending.Source = "SELECT Count(CommentID) as CommentsPendingCount FROM tblComment, tblBlog WHERE tblComment.commentInclude = 0 AND tblComment.blogID = tblBlog.blogID AND tblBlog.BlogAuthor = " & CInt(Session("MM_UserID"))
end if
rsComments_Pending.CursorType = 0
rsComments_Pending.CursorLocation = 2
rsComments_Pending.LockType = 1
rsComments_Pending.Open()
%>
<%
Dim rsComments
Dim rsComments_numRows
Set rsComments = Server.CreateObject("ADODB.Recordset")
rsComments.ActiveConnection = MM_blog_STRING
if Session("isAdmin") = 1 then
	rsComments.Source = "SELECT * FROM tblComment, tblBlog WHERE tblComment.commentInclude = 0 AND tblComment.blogID = tblBlog.blogID"
elseif Session("isAdmin") = 0 then
	rsComments.Source = "SELECT * FROM tblComment, tblBlog WHERE tblComment.commentInclude = 0 AND (tblComment.blogID = tblBlog.blogID) AND (tblBlog.BlogAuthor = " & CInt(Session("MM_UserID")) & ") ORDER BY commentDate ASC"
end if
rsComments.CursorType = 0
rsComments.CursorLocation = 2
rsComments.LockType = 1
rsComments.Open()
rsComments_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = -1
Repeat1__index = 0
rsComments_numRows = rsComments_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Approve Comments</title>
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
				<li><a href="approve_comments.asp">Comments (<%=(rsComments_Pending.Fields.Item("CommentsPendingCount").Value)%>)</a></li>
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



    <h2 align="left">Approve Comments </h2>
    <% If Not rsComments.EOF Or Not rsComments.BOF Then %>
<table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
<%
While ((Repeat1__numRows <> 0) AND (NOT rsComments.EOF))
%>
<tr>
<th><a href="<%=(rsComments.Fields.Item("commentURL").Value)%>"><%=(rsComments.Fields.Item("commentName").Value)%></a></th>
<th><a href="mailto:<%=(rsComments.Fields.Item("commentEmail").Value)%>" title="Email this user"><%=(rsComments.Fields.Item("commentEmail").Value)%></a></th>
<th><a href="confirm_publish.asp?passID=<%=(rsComments.Fields.Item("commentID").Value)%>">Approve</a>/<a href="edit_comment.asp?passID=<%=(rsComments.Fields.Item("commentID").Value)%>">Edit</a>/<a href="delete_comment.asp?passID=<%=(rsComments.Fields.Item("commentID").Value)%>">Delete</a> | <a href="template_permalink.asp?id=<%=(rsComments.Fields.Item("tblComment.BlogID").Value)%>" target="_blank">Original Post</a></th>
</tr>
<tr>
<td colspan="3"><%=(rsComments.Fields.Item("commentHTML").Value)%></td>
</tr>
<%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsComments.MoveNext()
Wend
%>
</table>
<% End If ' end Not rsComments.EOF Or NOT rsComments.BOF %>
		</div>
	</div>
</body>
</html>
<%
rsComments.Close()
Set rsComments = Nothing
%>

<%
rsComments_Pending.Close()
Set rsComments_Pending = Nothing
%>