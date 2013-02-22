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
  MM_editTable = "tblGallery"
  MM_editColumn = "fldGalleryID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "gallery.asp"
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
    MM_editCmd.ActiveConnection.Close
	
		
		'Variables to Set
		dim thisfilename
		thisfilename = "gallery_add.asp"
		dim filemanagerdir
		filemanagerdir = "\images\" 'Relative to where the root of the website is
		dim filemanagerdbdir
		filemanagerdbdir = MM_recordId
		
		'No need for setting parameters below
		
		
		
		Function CheckFolderExists(sFolderName)
			
			Dim FileSystemObject
			
			Set FileSystemObject = Server.CreateObject("Scripting.FileSystemObject")
			
			If (FileSystemObject.FolderExists(sFolderName)) Then
			CheckFolderExists = True
			Else
			CheckFolderExists = False
			End If
			
			Set FileSystemObject = Nothing
		
		End Function
		
		
		Foldertocreate = Server.MapPath(thisfilename) 
		if filemanagerdbdir = "" then
			Foldertocreate = Replace(Foldertocreate,thisfilename,(Right(filemanagerdir, Len(filemanagerdir)-1) & filemanagerdbdir))
		else
			Foldertocreate = Replace(Foldertocreate,thisfilename,(Right(filemanagerdir, Len(filemanagerdir)-1) & filemanagerdbdir & "\"))
		end if 
		If CheckFolderExists(Foldertocreate) Then
			'Response.Write(Foldertocreate)
			 'On Error Resume Next
			 Dim fso 
			 Set fso = Server.CreateObject("Scripting.FileSystemObject")
			 Set fso = fso.GetFolder(Foldertocreate)
			 fso.Delete
			 Set fso = Nothing
		Else
				'Response.Write("!")
		End If
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
<%
Dim rsGalleryDelete__MMColParam
rsGalleryDelete__MMColParam = "1"
If (Request.QueryString("fldGalleryID") <> "") Then 
  rsGalleryDelete__MMColParam = Request.QueryString("fldGalleryID")
End If
%>
<%
Dim rsGalleryDelete
Dim rsGalleryDelete_numRows
Set rsGalleryDelete = Server.CreateObject("ADODB.Recordset")
rsGalleryDelete.ActiveConnection = MM_blog_STRING
rsGalleryDelete.Source = "SELECT * FROM tblGallery WHERE fldGalleryID = " + Replace(rsGalleryDelete__MMColParam, "'", "''") + ""
rsGalleryDelete.CursorType = 0
rsGalleryDelete.CursorLocation = 2
rsGalleryDelete.LockType = 1
rsGalleryDelete.Open()
rsGalleryDelete_numRows = 0
%>
<%
curpath = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
curpath =  Left(curpath, InstrRev(curpath, "/"))
galleryroot = Right(curpath, Len(curpath) - Instr(curpath, "//")-1)
galleryroot = Right(galleryroot, Len(galleryroot) - Instr(galleryroot, "/")+1) & "images/"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Delete Gallery</title>
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
          <h2>Delete Gallery</h2>
          <form action="<%=MM_editAction%>" method="post" name="form1" id="form1"><table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
   <tr valign="baseline">
     <th nowrap="nowrap" align="right">Title:</th>
     <td><%=(rsGalleryDelete.Fields.Item("fldGalleryTitle").Value)%> </td>
   </tr>
   <tr class="tabledisplay">
     <th nowrap="nowrap" align="right" valign="top">Desc:</th>
     <td valign="baseline"><%=(rsGalleryDelete.Fields.Item("fldGalleryDesc").Value)%> </td>
   </tr>
   <tr valign="baseline">
     <th nowrap="nowrap" align="right">Pic:</th>
     <td><img src="thumbnailimage.aspx?filename=<%=galleryroot%><%=(rsGalleryDelete.Fields.Item("fldGalleryID").Value)%>/<%=(rsGalleryDelete.Fields.Item("fldGalleryPic").Value)%>&width=<%=(rsGalleryConfig.Fields.Item("fldGalleryTitleThumb").Value)%>" /></td>
   </tr>
   <tr valign="baseline">
     <td colspan="2" align="right" nowrap="nowrap"><div align="center">
         <input type="submit" value="Delete Gallery" />
     </div></td>
   </tr>
 </table>
    <input type="hidden" name="MM_delete" value="form1" />
    <input type="hidden" name="MM_recordId" value="<%= rsGalleryDelete.Fields.Item("fldGalleryID").Value %>" />
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
<%
rsGalleryDelete.Close()
Set rsGalleryDelete = Nothing
%>

