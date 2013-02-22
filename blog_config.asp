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
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_blog_STRING
    MM_editCmd.CommandText = "UPDATE tblBlogRSS SET blogTitle = ?, blogSubTitle = ?, blogDesc = ?, blogURL = ?, blogAuthor = ?, blogEmail = ?, blogPosts = ?, blogLayout = ? WHERE rssID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 255, Request.Form("txtTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("blogSubTitle")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 203, 1, 536870910, Request.Form("txtDesc")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 255, Request.Form("txtURL")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 255, Request.Form("txtAuthor")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 255, Request.Form("txtEmail")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("BlogPosts"), Request.Form("BlogPosts"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 202, 1, 50, Request.Form("blogLayout")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "blog_config.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>

<%
Dim rsConfig
Dim rsConfig_numRows
Set rsConfig = Server.CreateObject("ADODB.Recordset")
rsConfig.ActiveConnection = MM_blog_STRING
rsConfig.Source = "SELECT * FROM tblBlogRSS"
rsConfig.CursorType = 0
rsConfig.CursorLocation = 2
rsConfig.LockType = 1
rsConfig.Open()
rsConfig_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Blog Configuration</title>
<meta name="Description" content="" />
<meta name="Keywords" content="" />

	<style type="text/css" media="screen">	@import "tabs.css";	</style>
    <link rel="stylesheet" href="css/validationEngine.jquery.css" type="text/css" media="screen" title="no title" charset="utf-8" />
    <script src="js/jquery.min.js" type="text/javascript"></script>
    <script src="js/jquery.validationEngine-en.js" type="text/javascript"></script>
    <script src="js/jquery.validationEngine.js" type="text/javascript"></script>
        <script>	
		$(document).ready(function() {		
			$("#form1").validationEngine()
		});
	</script>
	
<script type="text/javascript">
function PreviewTheme()
	{
	var theme = document.getElementById('blogLayout').value;
	var path = 'default.asp?layout=' + theme;
	window.open(path);
	}
</script>
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
		<li><a href="users.asp">Users</a></li>
		<li><a href="layout.asp">Layout</a></li>
		<li><a class="current" href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">


   <h2>Blog Configuration</h2>
   <form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
<table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
<tr>
<th align="right" valign="middle">Title</th>
<td>
<input name="txtTitle" type="text" class="validate[required]" id="txtTitle" value="<%=(rsConfig.Fields.Item("blogTitle").Value)%>" size="40" />
<span class="req">*</span></td>
</tr>
<tr>
  <th align="right" valign="middle">SubTitle</th>
  <td><input name="blogSubTitle" type="text" class="txtBox" id="blogSubTitle" value="<%=(rsConfig.Fields.Item("blogSubTitle").Value)%>" size="40" /></td>
</tr>
<tr>
<th align="right" valign="middle">Description</th>
<td>
<textarea name="txtDesc" cols="40" rows="4" class="txtBox" id="txtDesc"><%=(rsConfig.Fields.Item("blogDesc").Value)%></textarea></td>
</tr>
<tr>
<th align="right" valign="middle">URL</th>
<td>
<input name="txtURL" type="text" class="validate[required]" id="txtURL" value="<%=(rsConfig.Fields.Item("blogURL").Value)%>" size="50" /> 
<span class="req">*</span> <strong>(should be "
<% 
    prot = "http" 
    https = lcase(request.ServerVariables("HTTPS")) 
    if https <> "off" then prot = "https" 
    domainname = Request.ServerVariables("SERVER_NAME") 
    filename = Request.ServerVariables("SCRIPT_NAME") 
	theurl = prot & "://" & domainname & filename
	thelen = Len(theurl)
	thelen = thelen - 15
	theurl = Left(theurl, thelen)
    response.write theurl
%>")</strong></td>
</tr>
<tr>
<th align="right" valign="middle">Author</th>
<td>
<input name="txtAuthor" type="text" class="txtBox" id="txtAuthor" value="<%=(rsConfig.Fields.Item("blogAuthor").Value)%>" size="40" /></td>
</tr>
<tr>
<th align="right" valign="middle">Email</th>
<td>
<input name="txtEmail" type="text" class="validate[custom[email]]" id="txtEmail" value="<%=(rsConfig.Fields.Item("blogEmail").Value)%>"  size="40" /></td>
</tr>
<tr>
  <th align="right" valign="middle"><span class="help" title="How many posts do you want to appear on you front default page?">Posts to show?</span> </th>
  <td><input name="BlogPosts" type="text" class="validate[required,custom[onlyNumber]]" id="BlogPosts" value="<%=(rsConfig.Fields.Item("BlogPosts").Value)%>" size="40" />
    <span class="req">*</span></td>
</tr>
<tr>
  <th align="right" valign="middle">Template </th>
  <td>    <select name="blogLayout" id="blogLayout">
<%
Dim objFSO
Dim objFolder
Dim objSubFolder
Dim objFile
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'Get Current Folder
Set objFolder = objFSO.GetFolder(Server.MapPath("Themes/"))
For Each objSubFolder In objFolder.SubFolders
	if CStr(objSubFolder.Name) = CStr(rsConfig.Fields.Item("blogLayout").Value) then
		extraText = "' selected='selected'"
	else
		extraText = "'"
	end if
    Response.Write "<option value='" & objSubFolder.Name & extraText &  ">" & objSubFolder.Name & "</option>" & vbcrlf
Next 
%>
    </select>
    <span class="req">*</span> <a href="javascript:void(PreviewTheme());">Preview</a> </td>
</tr>
<tr align="center" valign="middle">
<td colspan="2">
  <input name="Submit" type="submit" value="Update Configuration" /></td>
</tr>
</table>
<input type="hidden" name="MM_update" value="form1" />
<input type="hidden" name="MM_recordId" value="<%= rsConfig.Fields.Item("rssID").Value %>" />
</form>

   <h2>
     Additional Info </h2>
   <p><strong>IP Address:</strong>  <%= Request.ServerVariables("REMOTE_ADDR") %><br />
      <strong>Browser:</strong> <%= Request.ServerVariables("HTTP_USER_AGENT") %> </p>
   <TABLE border="0" align="center" cellpadding="0" cellspacing="1" class="tabledisplay">
<TR>
	<th>Server Variable</th>
	<th>Value</th>
</TR>
<% For Each Item In Request.ServerVariables %> 
<TR>
	<TD><%= Item %></TD>
	<TD><%= Request.ServerVariables(Item) %></TD>
</TR>
<% Next %>
</TABLE>

<p class="msgbox success">This is  success  <code>&lt;element class=&quot;msgbox success&quot;&gt;</code><strong> *element can be any tag </strong></p>
<p class="msgbox warning">This is warning <code>&lt;element class=&quot;msgbox warning&quot;&gt;</code></p>
<p class="msgbox error">This is  error <code>&lt;element class=&quot;msgbox error&quot;&gt;</code></p>
<p class="msgbox info">This is info  <code>&lt;element class=&quot;msgbox info&quot;&gt;</code></p>
		</div>
	</div>
</body>
</html>
<%
rsConfig.Close()
Set rsConfig = Nothing
%>