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
If request("Move") <> "" then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_blog_STRING
	MYSQL = "UPDATE tblBlog SET BlogCat = " & request("movecat") & " WHERE BlogCat = " & request("MM_recordId")
    MM_editCmd.CommandText = MYSQL
    MM_editCmd.Execute
	MYSQL = "Delete * FROM tblCat WHERE CatID = " & request("MM_recordId")
	  MM_editCmd.CommandText = MYSQL
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	response.Redirect("main.asp")
end if
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
  MM_editTable = "tblCat"
  MM_editColumn = "CatID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "cat.asp"
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
	MM_editQuery = "Delete * FROM tblBlog WHERE BlogCat = " & request("MM_recordId")
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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("CatID") <> "") Then
  Recordset1__MMColParam = Request.QueryString("CatID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_blog_STRING
Recordset1.Source = "SELECT * FROM tblCat WHERE CatID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
%>
<%
Dim rsPosts__MMColParam
rsPosts__MMColParam = "1"
If (Request.QueryString("Catid") <> "") Then 
  rsPosts__MMColParam = Request.QueryString("Catid")
End If
%>
<%
Dim rsPosts
Dim rsPosts_numRows
Set rsPosts = Server.CreateObject("ADODB.Recordset")
rsPosts.ActiveConnection = MM_blog_STRING
rsPosts.Source = "SELECT * FROM tblBlog WHERE BlogCat = " + Replace(rsPosts__MMColParam, "'", "''") + " ORDER BY BlogHeadline ASC"
rsPosts.CursorType = 0
rsPosts.CursorLocation = 2
rsPosts.LockType = 1
rsPosts.Open()
rsPosts_numRows = 0
%>
<%
Dim rsCats__MMColParam
rsCats__MMColParam = "1"
If (Request.QueryString("CatID") <> "") Then 
  rsCats__MMColParam = Request.QueryString("CatID")
End If
%>
<%
Dim rsCats
Dim rsCats_numRows
Set rsCats = Server.CreateObject("ADODB.Recordset")
rsCats.ActiveConnection = MM_blog_STRING
rsCats.Source = "SELECT * FROM tblCat WHERE CatID <> " + Replace(rsCats__MMColParam, "'", "''") + " ORDER BY CatName ASC"
rsCats.CursorType = 0
rsCats.CursorLocation = 2
rsCats.LockType = 1
rsCats.Open()
rsCats_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = -1
Repeat1__index = 0
rsPosts_numRows = rsPosts_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Delete a Category</title>
	<style type="text/css" media="screen">
	@import "tabs.css";
	</style>
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
		<li><a class="current" href="cat.asp">Categories</a></li>
			<ul id="secondary">
				<li><a href="cat_add.asp">Add a Category</a></li>
			</ul>		
		<li><a href="users.asp">Users</a></li>
		<li><a href="layout.asp">Layout</a></li>
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">
<h2>Delete a Category</h2>
        <form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
          <h3>Delete Category: <strong><%=(Recordset1.Fields.Item("CatName").Value)%></strong></h3>
          <p>Warning: This could potentially delete posts related to this category:</p>
          <p><% If not rsPosts.EOF Or not rsPosts.BOF Then %>
          <%  noposts = 0
While ((Repeat1__numRows <> 0) AND (NOT rsPosts.EOF)) 
%>      
            <a href="template_archives_cat.asp?cat=<%=(rsPosts.Fields.Item("BlogCat").Value)%>#<%=(rsPosts.Fields.Item("BlogID").Value)%>" target="_blank" title="View this post in a new window"><%=(rsPosts.Fields.Item("BlogHeadline").Value)%></a>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsPosts.MoveNext() %>
  <br /> <%
Wend
%></p>
<% Else 
	noposts = 1 %>
			<p>No related posts found for this category.</p>
			<% End If ' end Not rsPosts.EOF Or NOT rsPosts.BOF %>
<p>
            <input type="submit" name="Submit" value="Delete!" <% if rsCats.EOF Then response.Write("disabled='disabled'") end if %> />
            <% if NOT rsCats.EOF AND noposts <> 1 Then %> 
          or</p>
<p>
  <input type="submit" name="Move" value="Move!" /> 
			all posts above to Category
			<label>
			<select name="movecat" id="movecat">
			  <%
While (NOT rsCats.EOF)
%>
			  <option value="<%=(rsCats.Fields.Item("CatID").Value)%>"><%=(rsCats.Fields.Item("CatName").Value)%></option>
			  <%
  rsCats.MoveNext()
Wend
If (rsCats.CursorType > 0) Then
  rsCats.MoveFirst
Else
  rsCats.Requery
End If
%>
		    </select>
			</label> 
		  and also <strong>delete the current category</strong>
          <% elseif noposts <> 1 then %> 
		  </p>
<p>No other categories available to move all posts to therefore the delete is disabled b/c you must have at least one category. If you'd like to move all of the above posts to a new category before deleting this category then first <a href="cat_add.asp">create a new category</a> and return to this screen. Of course, since there are no other categories available to move these posts to you can simply <a href="cat_update.asp?catid=<%=(Recordset1.Fields.Item("CatID").Value)%>">rename this category</a>.
		  <% end if %>
</p>
<p>Note: This cannot be undone! </p>
          <input type="hidden" name="MM_delete" value="form1" />
          <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("CatID").Value %>" />
        </form>
    
		</div>
	</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
rsPosts.Close()
Set rsPosts = Nothing
%>
<%
rsCats.Close()
Set rsCats = Nothing
%>

