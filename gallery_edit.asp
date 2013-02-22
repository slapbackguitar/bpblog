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
Dim rsGalleryEdit__MMColParam
rsGalleryEdit__MMColParam = "1"
If (Request.QueryString("fldGalleryID") <> "") Then 
  rsGalleryEdit__MMColParam = Request.QueryString("fldGalleryID")
End If
%>
<%
Dim rsGalleryEdit
Dim rsGalleryEdit_numRows
Set rsGalleryEdit = Server.CreateObject("ADODB.Recordset")
rsGalleryEdit.ActiveConnection = MM_blog_STRING
rsGalleryEdit.Source = "SELECT * FROM tblGallery WHERE fldGalleryID = " + Replace(rsGalleryEdit__MMColParam, "'", "''") + ""
rsGalleryEdit.CursorType = 0
rsGalleryEdit.CursorLocation = 2
rsGalleryEdit.LockType = 1
rsGalleryEdit.Open()
rsGalleryEdit_numRows = 0
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
' *** Update Record: set variables
If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then
  MM_editConnection = MM_blog_STRING
  MM_editTable = "tblGallery"
  MM_editColumn = "fldGalleryID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "gallery.asp"
  MM_fieldsStr  = "fldGalleryTitle|value|fldGalleryDesc|value|fldGalleryPic|value"
  MM_columnsStr = "fldGalleryTitle|',none,''|fldGalleryDesc|',none,''|fldGalleryPic|',none,''"
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
<%
'Variables to Set
dim thisfilename
thisfilename = "gallery_edit.asp"
dim filemanagerdir
filemanagerdir = "\images\" 'Relative to where the root of the website is
dim filemanagerdbdir
filemanagerdbdir = (rsGalleryEdit.Fields.Item("fldGalleryID").Value)
dim tableclass
tableclass = "tabledisplay"
dim filemanagerthumbnailsize
filemanagerthumbnailsize = (rsGalleryConfig.Fields.Item("fldGalleryThumb").Value)
'No need for setting parameters below
%>
<%
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
%> 
<%
Foldertocreate = Server.MapPath(thisfilename) 
if filemanagerdbdir = "" then
	Foldertocreate = Replace(Foldertocreate,thisfilename,(Right(filemanagerdir, Len(filemanagerdir)-1) & filemanagerdbdir))
else
	Foldertocreate = Replace(Foldertocreate,thisfilename,(Right(filemanagerdir, Len(filemanagerdir)-1) & filemanagerdbdir & "\"))
end if 
If CheckFolderExists(Foldertocreate) Then
	'Response.Write("!")
Else
	Set fs = CreateObject("Scripting.FileSystemObject") 
	Set a = fs.CreateFolder(Foldertocreate)  
	Set fs=nothing
End If

if Request("filetodelete") <> "" then
	filetodelete2 = Request("filetodelete") 
	filetodelete = Server.MapPath(thisfilename) 
	if filemanagerdbdir = "" then
		filetodelete = Replace(filetodelete,thisfilename,(Right(filemanagerdir, Len(filemanagerdir)-1) & filemanagerdbdir) & filetodelete2)
	else
		filetodelete = Replace(filetodelete,thisfilename,(Right(filemanagerdir, Len(filemanagerdir)-1) & filemanagerdbdir & "\" & filetodelete2))
	end if 
	
	'Response.Write(filetodelete)
	Dim objFSOdel
	Set objFSOdel = Server.CreateObject("Scripting.FileSystemObject")
	objFSOdel.DeleteFile filetodelete, True
	Set objFSOdel = Nothing 
 End If
%> 
<%
curpath = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
curpath =  Left(curpath, InstrRev(curpath, "/"))
galleryroot = Right(curpath, Len(curpath) - Instr(curpath, "//")-1)
galleryroot = Right(galleryroot, Len(galleryroot) - Instr(galleryroot, "/")+1) & "images/"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Edit a Gallery</title>

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
          <h2>Edit a Gallery</h2>
<h4>Be sure to add a few photos if none are available before filling in the Title and Description and clicking &quot;Update Gallery&quot; b/c the uploads are processed first.</h4>
 <form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
   <table width="99%" border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
     <tr valign="baseline">
       <th width="10%" align="right" nowrap="nowrap">Title:</th>
       <td width="90%"><input type="text" name="fldGalleryTitle" id="fldGalleryTitle" value="<%=(rsGalleryEdit.Fields.Item("fldGalleryTitle").Value)%>" size="32" class="validate[required]" /> <span class="req">*</span>       </td>
     </tr>
     <tr>
       <th nowrap="nowrap" align="right" valign="top">Desc:</th>
       <td valign="baseline"><textarea name="fldGalleryDesc" cols="40" rows="5"><%=(rsGalleryEdit.Fields.Item("fldGalleryDesc").Value)%></textarea></td>
     </tr>
     <tr valign="baseline" class="tablenormal">
       <th nowrap="nowrap" align="right">Current Pic: </th>
       <td><% if (rsGalleryEdit.Fields.Item("fldGalleryPic").Value) <> "" Then %><img src="thumbnailimage.aspx?filename=<%=galleryroot%><%=(rsGalleryEdit.Fields.Item("fldGalleryID").Value)%>/<%=(rsGalleryEdit.Fields.Item("fldGalleryPic").Value)%>&width=<%=(rsGalleryConfig.Fields.Item("fldGalleryTitleThumb").Value)%>" class="thumbnail" /><% else %> None Selected Yet <% end if %></td>
     </tr>
     <tr valign="baseline">
       <th nowrap="nowrap" align="right">Pic:</th>
       <td><select name="fldGalleryPic" id="fldGalleryPic"><% ListFolderContents(Server.MapPath(galleryroot & (rsGalleryEdit.Fields.Item("fldGalleryID").Value) & "/")) %>
<% sub ListFolderContents(path)
     dim fs, folder, file, item, url
     set fs = CreateObject("Scripting.FileSystemObject")
     set folder = fs.GetFolder(path)
    'Display the target folder and info.
     'Response.Write("<li><b>" & folder.Name & "</b> - " _
       '& folder.Files.Count & " files, ")
     if folder.SubFolders.Count > 0 then
       %><%
     end if
     'Response.Write(Round(folder.Size / 1024) & " KB total." _
       '& vbCrLf)
     'Response.Write("<ul>" & vbCrLf)
     'Display a list of sub folders.
     for each item in folder.SubFolders
       ListFolderContents(item.Path)
     next
	 
	 	   if folder.Files.Count <> "" then %>
	   <% end if
     'Display a list of files.
       %>  
	     <%
     for each item in folder.Files
       url = MapURL(item.path)
	   if (right(item.Name, 3) = "jpg" OR right(item.Name, 3) = "JPG" OR right(item.Name, 3) = "PNG" OR right(item.Name, 3) = "png"  OR right(item.Name, 4) = "jpeg" OR right(item.Name, 3) = "JPEG" OR right(item.Name, 3) = "gif" OR right(item.Name, 3) = "GIF") then
       %>
<option value="<%=item.name%>" <% if item.name = (rsGalleryEdit.Fields.Item("fldGalleryPic").Value) then%>selected<%end if%>><%=item.name%></option>
	   
	   <%
	   end if
     next %>        <%
     'Response.Write("</ul>" & vbCrLf)
     'Response.Write("</li>" & vbCrLf)
   end sub
   function MapURL(path)
     dim rootPath, url
     'Convert a physical file path to a URL for hypertext links.
     rootPath = Server.MapPath("/")
     url = Right(path, Len(path) - Len(rootPath))
     MapURL = Replace(url, "\", "/")
   end function %></select> 
       </td>
     </tr>
     <tr align="center" valign="middle">
       <td colspan="2" align="left" nowrap="nowrap"><input type="submit" value="Update Gallery" />       </td>
       </tr>
   </table>
   
<input type="hidden" name="MM_update" value="form1" />
<input type="hidden" name="MM_recordId" value="<%= rsGalleryEdit.Fields.Item("fldGalleryID").Value %>" />
 </form>
 <%
Path = galleryroot & (rsGalleryEdit.Fields.Item("fldGalleryID").Value) & "/"
Session("path") = Path
%>
<%
thumbsize = (rsGalleryConfig.Fields.Item("fldGalleryThumb").Value)
Set fso = Server.CreateObject("Scripting.FileSystemObject")	
If Right(Path,1)="/" AND Path<>"/" Then Path=Left(Path,Len(Path)-1)
'response.write("<font color='white'><b>" & Path & "</font></b><br>")
Var =InstrRev(Path,"/")
dirup=left(Path,Var)
'response.write ("[<a href='browser.asp?path=" & dirup & "'>Directory up</a>]") 
%>
<% 
aktion=request.querystring("aktion")
Set ts=fso.GetFolder(Server.MapPath(Path))
Select Case aktion
	Case "deletefile"	
		fso.DeleteFile(Server.MapPath(request.querystring("file")))
		redirecturl = "gallery_edit.asp?fldGalleryID=" & Request("fldGalleryID")
		response.redirect(redirecturl)
End select
%>
	<table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
<%
If Path<>"/" AND Right(Path,1)<>"/" then Path=Path & "/"
Pos=instr(right(Path,Len(Path)-1),"/")
If Path="/" then Pos=0
If int(Pos)=0 then '->If Path is Root-Directory
	FirstFolder="/" 
	ShowFiles=false
	ShowFolders=true
	ShowUpload=false		
Else 
	FirstFolder=right(Path,Len(Path)-1)
	If FirstFolder<>"" then FirstFolder=Left(FirstFolder,Pos-1)
	If Instr(Session("aspEdit_FolderAccess"),"," & FirstFolder & ",")>0 OR Session("aspEdit_Level")=3 then 
		ShowFiles=true
		ShowFolders=true
		ShowUpload=true
	else
		ShowFiles=true
		ShowFolders=true
		ShowUpload=true
	end if
End If
If ShowFolders=true then
For each SubF in ts.Subfolders
If right(Path,1)="/" then
	WholeSubF=Path & SubF.Name
else
	WholeSubF=Path & "/" & SubF.Name
end if
If (Path="/" AND Instr(Session("aspEdit_FolderAccess"),"," & SubF.Name & ",")>0) OR (ShowFolders=true AND Path<>"/") OR (Session("aspEdit_Level")=3) then
	ShowThisFolder=true
else
	ShowThisFolder=true
end if
If ShowThisFolder=true then
%>
<%
End If
next
End If
%>
		<tr>
			<th colspan="7">Files
			</th>
		</tr>
<%
If ShowFiles=true then
For each File in ts.files
If right(Path,1)="/" then
	WholeFile=Path & File.Name
else
	WholeFile=Path & "/" & File.Name
end if
Var=InstrRev(File.Name,".")
FileType=Right(File.Name,Len(File.Name)-Var)
%>
		<tr>
			<td><% if right(File.Name, 3) = "jpg" OR right(File.Name, 3) = "JPG" OR right(File.Name, 3) = "PNG" OR right(File.Name, 3) = "png"  OR right(File.Name, 4) = "jpeg" OR right(File.Name, 3) = "JPEG" OR right(File.Name, 3) = "gif" OR right(File.Name, 3) = "GIF" then %>
			<img src="thumbnailimage.aspx?filename=<%=Path%><%=File.Name%>&width=<%=thumbsize%>" alt="<%=File.Name%>" class="thumbnail" /> <%else%> <%=File.Name%> <% end if %>
			</td>
			<td>
			    <%=File.Type%></td>
			<td>  <%
if File.Size <1024 Then
    Response.Write File.Size & " B"
ElseIf File.Size < 1048576 Then
    Response.Write Round(File.Size / 1024.1) & " KB"
Else
    Response.Write Round((File.Size/1024)/1024.1) & " MB"
End if
Var=InstrRev(File.Name,".")
FileType=Right(File.Name,Len(File.Name)-Var)
%>
		  </td>
			<td>			  <%Response.Write File.DateLastModified%></td><td>				  <a href="gallery_edit.asp?fldGalleryID=<%=(rsGalleryEdit.Fields.Item("fldGalleryID").Value)%>&amp;path=<%=path & "&amp;aktion=deletefile&amp;file=" & WholeFile%>">delete</a></td>
		</tr>
<%
next
End If
%>
	</table>
<form method="post" enctype="multipart/form-data" action="upload.asp"> 
  
  <input type="file" size="50" name="FILE1" />
  <input type="submit" value="Upload!" />
  <input name="path" type="hidden" value="<%=path%>" />
  
</form>
<%
	Set fso = Nothing
	Set ts = Nothing
%>
 
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
rsGalleryEdit.Close()
Set rsGalleryEdit = Nothing
%>

