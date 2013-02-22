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
  MM_editTable = "tblLayout"
  MM_editColumn = "layoutid"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "layout.asp"
  MM_fieldsStr  = "layoutTitle|value|layout1|value|layout2|value|layout3|value|layout4|value|layout5|value"
  MM_columnsStr = "layoutTitle|',none,''|layout1|',none,''|layout2|',none,''|layout3|',none,''|layout4|',none,''|layout5|',none,''"
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
Dim rsLayouts__MMColParam
rsLayouts__MMColParam = "1"
If (Request.QueryString("layoutid") <> "") Then 
  rsLayouts__MMColParam = Request.QueryString("layoutid")
End If
%>
<%
Dim rsLayouts
Dim rsLayouts_numRows
Set rsLayouts = Server.CreateObject("ADODB.Recordset")
rsLayouts.ActiveConnection = MM_blog_STRING
rsLayouts.Source = "SELECT * FROM tblLayout WHERE layoutid = " + Replace(rsLayouts__MMColParam, "'", "''") + ""
rsLayouts.CursorType = 0
rsLayouts.CursorLocation = 2
rsLayouts.LockType = 1
rsLayouts.Open()
rsLayouts_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Update a Layout</title>
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
		<li><a href="users.asp">Users</a></li>
		<li><a class="current" href="layout.asp">Layout</a></li>
			<ul id="secondary">
				<li><a href="layout_add.asp">Add a New Layout</a></li>
			</ul>		
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">
          <h2>Update a Layout</h2>
        <form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
          <table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" class="tabledisplay">
            <tr>
              <th width="10%" align="right" valign="top" nowrap="nowrap">Title:</th>
              <td width="90%" valign="baseline"><label>
                <input name="layoutTitle" type="text" id="layoutTitle" value="<%=(rsLayouts.Fields.Item("layoutTitle").Value)%>" size="40" maxlength="100" />
              </label></td>
            </tr>
            <tr>
              <th nowrap="nowrap" align="right" valign="top">Layout1:</th>
              <td valign="baseline"><textarea name="layout1" cols="50" rows="8"><%=(rsLayouts.Fields.Item("layout1").Value)%></textarea>              </td>
            </tr>
            <tr>
              <th nowrap="nowrap" align="right" valign="top">Layout2:</th>
              <td valign="baseline"><textarea name="layout2" cols="50" rows="8"><%=(rsLayouts.Fields.Item("layout2").Value)%></textarea>              </td>
            </tr>
            <tr>
              <th nowrap="nowrap" align="right" valign="top">Layout3:</th>
              <td valign="baseline"><textarea name="layout3" cols="50" rows="8"><%=(rsLayouts.Fields.Item("layout3").Value)%></textarea>              </td>
            </tr>
            <tr>
              <th nowrap="nowrap" align="right" valign="top">Layout4:</th>
              <td valign="baseline"><textarea name="layout4" cols="50" rows="8"><%=(rsLayouts.Fields.Item("layout4").Value)%></textarea>              </td>
            </tr>
            <tr>
              <th nowrap="nowrap" align="right" valign="top">Layout5:</th>
              <td valign="baseline"><textarea name="layout5" cols="50" rows="8"><%=(rsLayouts.Fields.Item("layout5").Value)%></textarea>              </td>
            </tr>
            <tr valign="baseline">
              <td colspan="2" align="left" valign="middle" nowrap="nowrap"><input name="submit" type="submit" value="Update Layout" />              </td>
            </tr>
          </table>
      
          <input type="hidden" name="MM_update" value="form1">
          <input type="hidden" name="MM_recordId" value="<%= rsLayouts.Fields.Item("layoutid").Value %>">
</form>
		</div>
	</div>
</body>
</html>
<%
rsLayouts.Close()
Set rsLayouts = Nothing
%>

