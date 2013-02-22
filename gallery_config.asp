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
' *** Update Record: set variables
If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then
  MM_editConnection = MM_blog_STRING
  MM_editTable = "tblGalleryConfig"
  MM_editColumn = "fldGalleryConfigID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "gallery.asp"
  MM_fieldsStr  = "fldGalleryThumb|value|fldGalleryTitleThumb|value"
  MM_columnsStr = "fldGalleryThumb|none,none,NULL|fldGalleryTitleThumb|none,none,NULL"
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
' *** Update Record: construct a sql update statement and execute it
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then
  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId
  If (Not MM_abortEdit) Then
    ' execute the update
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
<%
Dim rsGalleryConfig
Dim rsGalleryConfig_numRows
Set rsGalleryConfig = Server.CreateObject("ADODB.Recordset")
rsGalleryConfig.ActiveConnection = MM_blog_STRING
rsGalleryConfig.Source = "SELECT * FROM tblGalleryConfig"
rsGalleryConfig.CursorType = 0
rsGalleryConfig.CursorLocation = 2
rsGalleryConfig.LockType = 1
rsGalleryConfig.Open()
rsGalleryConfig_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Gallery Configuration</title>

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
		<li><a class="current" href="gallery.asp">Gallery</a></li>
			<ul id="secondary">
				<li><a href="gallery_add.asp">Create a New Gallery</a></li>
			</ul>
		<% if Session("isAdmin") = 1 then %>
		<li><a href="pages.asp">Pages</a></li>
		<li><a href="cat.asp">Categories</a></li>
		<li><a href="users.asp">Users</a></li>
		<li><a href="layout.asp">Layout</a></li>
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>	<div id="main">
		<div id="contents">
          <h2>Gallery Configuration</h2>
 <form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
   <table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
        <tr valign="baseline">
       <th nowrap="nowrap" align="right">Gallery Title Thumbnail Size (width in pixels):</th>
       <td><input type="text" name="fldGalleryTitleThumb" id="fldGalleryTitleThumb" value="<%=(rsGalleryConfig.Fields.Item("fldGalleryTitleThumb").Value)%>" size="32" class="validate[required,custom[onlyNumber]]" />
       </td>
     </tr>
     <tr valign="baseline">
       <th nowrap="nowrap" align="right">Gallery Thumbnail Size (width in pixels):</th>
       <td><input type="text" name="fldGalleryThumb" id="fldGalleryThumb" value="<%=(rsGalleryConfig.Fields.Item("fldGalleryThumb").Value)%>" size="32" class="validate[required,custom[onlyNumber]]" />
       </td>
     </tr>
     <tr align="center" valign="middle">
       <td colspan="2" nowrap="nowrap"><input type="submit" value="Update Configuration" />       </td>
       </tr>
   </table>
   <input type="hidden" name="MM_update" value="form1" />
   <input type="hidden" name="MM_recordId" value="<%= rsGalleryConfig.Fields.Item("fldGalleryConfigID").Value %>" />
 </form>
 
		</div>
	</div>
</body>
</html>
<%
rsConfig.Close()
Set rsConfig = Nothing
%>
<%
rsGalleryConfig.Close()
Set rsGalleryConfig = Nothing
%>

