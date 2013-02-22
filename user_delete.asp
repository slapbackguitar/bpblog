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
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If
' boolean to abort record edit
MM_abortEdit = false
' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables
if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then
  MM_editConnection = MM_blog_STRING
  MM_editTable = "tblAuthor"
  MM_editColumn = "fldAuthorID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "users.asp"

End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it
If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then
  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId
  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
	MM_editCmd.CommandText = "delete from tblBlog where BlogAuthor = " & MM_recordId
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If
End If
%>
<%
Dim rsUsers__MMColParam
rsUsers__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rsUsers__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rsUsers
Dim rsUsers_numRows
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.ActiveConnection = MM_blog_STRING
rsUsers.Source = "SELECT * FROM tblAuthor WHERE fldAuthorID = " + Replace(rsUsers__MMColParam, "'", "''") + ""
rsUsers.CursorType = 0
rsUsers.CursorLocation = 2
rsUsers.LockType = 1
rsUsers.Open()
rsUsers_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Delete a User</title>
	<style type="text/css" media="screen">@import "tabs.css";</style>
</head>
<body>
	<% if Session("MM_Admin") = 1 then %>
	<h3 class="floatright"><a href="?view=1" accesskey="2">User View</a> | <a href="?view=2" accesskey="3">Admin View</a></h3>
	<% end if %>
	<h1><a href="main.asp" accesskey="1">bp blog admin (<%=Session("MM_Username")%>)</a> | <a href="default.asp">Your Blog</a></h1>
	<div id="header">
	<ul id="primary">
		<li><a href="main.asp">Home (Entries)</a></li>
		<li><a href="user_update.asp?id=<%=Session("MM_UserID")%>">Profile</a></li>
		<li><a href="gallery.asp">Gallery</a></li>
		<% if Session("isAdmin") = 1 then %>
		<li><a href="pages.asp">Pages</a></li>
		<li><a href="cat.asp">Categories</a></li>
		<li><a  class="current"href="users.asp">Users</a></li>
			<ul id="secondary">
				<li><a href="user_add.asp">Add a New User</a></li>
			</ul>		
		<li><a href="layout.asp">Layout</a></li>
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">
          <h2>Delete a User</h2>
      <p align="left"><strong>Username:</strong>  <%=(rsUsers.Fields.Item("fldAuthorUsername").Value)%></p>
      <p align="left" style="font-weight: bold">Warning: This will delete this user's posts also! </p>
      <form name="form1" id="form1" method="post" action="<%=MM_editAction%>">
        <input type="submit" name="Submit" value="Delete this user" />
        <input type="hidden" name="MM_delete" value="form1" />
        <input type="hidden" name="MM_recordId" value="<%= rsUsers.Fields.Item("fldAuthorID").Value %>" />
      </form>
 
		</div>
	</div>
</body>
</html>
<%
rsUsers.Close()
Set rsUsers = Nothing
%>
