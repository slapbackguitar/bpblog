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
	rsComments_Pending.Source = "SELECT Count(*) as CommentsPendingCount FROM tblComment, tblBlog WHERE tblComment.commentInclude = 0 AND (tblComment.blogID = tblBlog.blogID) AND (tblBlog.BlogAuthor = " & CInt(Session("MM_UserID")) & ")"
end if
rsComments_Pending.CursorType = 0
rsComments_Pending.CursorLocation = 2
rsComments_Pending.LockType = 1
rsComments_Pending.Open()
%>
<%
' *** Edit Operations: declare variables
Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd
Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId
Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If
' boolean to abort record edit
MM_abortEdit = false
' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables
If (CStr(Request("MM_insert")) = "fHtmlEditor") Then
  MM_editConnection = MM_blog_STRING
  MM_editTable = "tblBlog"
  MM_editRedirectUrl = "rss.asp"
  MM_fieldsStr  = "txtHeading|value|BlogHTML|value|cat|value|BlogCommentInclude|value|BlogReadMore|value|BlogDraft|value|BlogAuthor|value"
  MM_columnsStr = "BlogHeadline|',none,''|BlogHTML|',none,''|BlogCat|none,none,NULL|BlogCommentInclude|none,none,NULL|BlogReadMore|none,none,NULL|BlogDraft|none,none,NULL|BlogAuthor|none,none,NULL"
  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next
  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it
Dim MM_tableValues
Dim MM_dbValues
If (CStr(Request("MM_insert")) <> "") Then
  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"
  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If
End If
%>
<%
Dim rs_cat
Dim rs_cat_numRows
Set rs_cat = Server.CreateObject("ADODB.Recordset")
rs_cat.ActiveConnection = MM_blog_STRING
rs_cat.Source = "SELECT *  FROM tblCat  ORDER BY CatName ASC"
rs_cat.CursorType = 0
rs_cat.CursorLocation = 2
rs_cat.LockType = 1
rs_cat.Open()
rs_cat_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Add Blog</title>
	<style type="text/css" media="screen">@import "tabs.css";</style>
    <link rel="stylesheet" href="css/validationEngine.jquery.css" type="text/css" media="screen" title="no title" charset="utf-8" />
    <script src="js/jquery.min.js" type="text/javascript"></script>
    <script src="js/jquery.validationEngine-en.js" type="text/javascript"></script>
    <script src="js/jquery.validationEngine.js" type="text/javascript"></script>
    		<script>	
		$(document).ready(function() {		
			$("#fHtmlEditor").validationEngine()
		});
	</script>
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
          <h2>Add Blog</h2>
          <form action="<%=MM_editAction%>" method="POST" name="fHtmlEditor" id="fHtmlEditor">
  <table width="99%" border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
    <tr>
      <th align="right">Blog Headline</th>
      <td><input name="txtHeading" type="text" class="validate[required]" id="txtHeading" size="50" /> 
        <span class="req">*</span> </td>
    </tr>
    <tr>
      <th width="9%" align="right">Category </th>
      <td width="91%"><select name="cat" id="cat">
        <%
While (NOT rs_cat.EOF)
%>
        <option value="<%=(rs_cat.Fields.Item("CatID").Value)%>"><%=(rs_cat.Fields.Item("CatName").Value)%></option>
        <%
  rs_cat.MoveNext()
Wend
If (rs_cat.CursorType > 0) Then
  rs_cat.MoveFirst
Else
  rs_cat.Requery
End If
%>
      </select></td>
    </tr>
    <tr>
      <th align="right">Comments </th>
      <td><strong>
        <select name="BlogCommentInclude" id="BlogCommentInclude">
          <option value="1" selected="selected">Yes</option>
          <option value="0">No</option>
        </select>
      </strong></td>
    </tr>
    <tr>
      <th align="right">Read More</th>
      <td><strong>
        <select name="BlogReadMore" id="BlogReadMore">
          <option value="0" selected="selected">No</option>
          <option value="1">Yes</option>
        </select>
      </strong></td>
    </tr>
    <tr>
      <th align="right">Draft</th>
      <td><strong>
        <select name="BlogDraft" id="BlogDraft">
          <option value="0" selected="selected">No</option>
          <option value="1">Yes</option>
        </select>
      </strong></td>
    </tr>
  </table>
  <input name="BlogAuthor" type="hidden" id="BlogAuthor" value="<%=Session("MM_UserID")%>" />
<!-- #INCLUDE file="ckeditor/ckeditor.asp" -->
<!-- #INCLUDE file="ckfinder/ckfinder.asp" -->
<%
dim editor
set editor = New CKEditor
editor.basePath = theBasePath
CKFinder_SetupCKEditor editor, replace(theConfigUserFilesPath,"UserFiles","ckfinder"), empty, empty
editor.editor "BlogHTML", initialValue
%>
  <input name="Submit" type="submit" value="Add Blog Entry" />
  <input type="hidden" name="MM_insert" value="fHtmlEditor" />
</form>
		</div>
	</div>
</body>
</html>
<%
rs_cat.Close()
Set rs_cat = Nothing
%>
<%
rsComments_Pending.Close()
Set rsComments_Pending = Nothing
%>