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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_blog_STRING
    MM_editCmd.CommandText = "INSERT INTO tblAuthor (fldAuthorUsername, fldAuthorPassword, fldAuthorRealName, fldAuthorEmail, fldAuthorWebsite, Approved, fldAdmin, fldAuthorBlurb) VALUES (?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 100, Request.Form("fldAuthorUsername")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("fldAuthorPassword")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 100, Request.Form("fldAuthorRealName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 100, Request.Form("fldAuthorEmail")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 100, Request.Form("fldAuthorWebsite")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("Approved"), Request.Form("Approved"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("fldAdmin"), Request.Form("fldAdmin"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 203, 1, 536870910, Request.Form("fldAuthorBlurb")) ' adLongVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "users.asp"
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
Dim rsUsers
Dim rsUsers_numRows
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.ActiveConnection = MM_blog_STRING
rsUsers.Source = "SELECT * FROM tblAuthor ORDER BY fldAuthorID ASC"
rsUsers.CursorType = 0
rsUsers.CursorLocation = 2
rsUsers.LockType = 1
rsUsers.Open()
rsUsers_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Add a New User</title>

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
          <h2>Add a User</h2>
   <form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
     <table width="99%" class="tabledisplay">
       <tr valign="baseline">
         <th width="10%" align="right" nowrap="nowrap">Username:</th>
         <td width="90%"><input type="text" name="fldAuthorUsername" id="fldAuthorUsername" value="" size="32" class="validate[required,custom[noSpecialCaracters]]" /> <span class="req">*</span> </td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Password:</th>
         <td><input type="password" name="fldAuthorPassword" id="fldAuthorPassword" value="" size="32" class="validate[required]" /> <span class="req">*</span> </td>
       </tr>
       <tr valign="baseline">
         <th align="right">Password Confirm:</th>
         <td><input name="fldAuthorPasswordConfirm" type="password" id="fldAuthorPasswordConfirm" size="32" class="validate[required,confirm[fldAuthorPassword]]" />
           <span class="req">*</span></td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Real Name:</th>
         <td><input type="text" name="fldAuthorRealName" id="fldAuthorRealName" value="" size="32" class="validate[required]" /> <span class="req">*</span> </td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Email:</th>
         <td><input type="text" name="fldAuthorEmail" id="fldAuthorEmail" value="" size="32" class="validate[required,custom[email]]" /> <span class="req">*</span> </td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Website:</th>
         <td><input type="text" name="fldAuthorWebsite" value="" size="32" />         </td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Approved:</th>
         <td><label>
         <select name="Approved" id="Approved">
           <option value="1" selected="selected">Yes</option>
           <option value="0">No</option>
         </select>
         </label></td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Admin:</th>
         <td><label>
           <select name="fldAdmin" id="fldAdmin">
             <option value="1">Yes</option>
             <option value="0" selected="selected">No</option>
                      </select>
         </label></td>
       </tr>
       <tr valign="baseline">
         <th nowrap="nowrap" align="right">Blurb:</th>
         <td> <!-- #INCLUDE file="ckeditor/ckeditor.asp" -->
<!-- #INCLUDE file="ckfinder/ckfinder.asp" -->
<%
dim editor
set editor = New CKEditor
editor.basePath = theBasePath
CKFinder_SetupCKEditor editor, replace(theConfigUserFilesPath,"UserFiles","ckfinder"), empty, empty
editor.editor "fldAuthorBlurb", initialValue
%></td>
       <tr align="center" valign="middle">
         <td colspan="2" align="left" nowrap="nowrap"><input type="submit" value="Add User" />         </td>
        </tr>
     </table>
     <input type="hidden" name="MM_insert" value="form1" />
   </form>
   <p>&nbsp;</p>
 </h3>
      
		</div>
	</div>
</body>
</html>
<%
rsUsers.Close()
Set rsUsers = Nothing
%>

