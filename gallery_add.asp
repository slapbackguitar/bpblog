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
Set MM_editCmd = Server.CreateObject("ADODB.Command")
MM_editCmd.ActiveConnection = MM_blog_STRING
MM_editCmd.CommandText = "insert into tblGallery (fldGalleryTitle, fldGalleryUser) VALUES ('Untitled', " & Session("MM_UserID")  & ")"
MM_editCmd.Execute
MM_editCmd.ActiveConnection.Close

Dim rsGalleryAdd

Set rsGalleryAdd = Server.CreateObject("ADODB.Recordset")
rsGalleryAdd.ActiveConnection = MM_blog_STRING
rsGalleryAdd.Source = "SELECT * FROM tblGallery WHERE fldGalleryID = (Select max(fldGalleryID)  from tblGallery)"
rsGalleryAdd.CursorType = 0
rsGalleryAdd.CursorLocation = 2
rsGalleryAdd.LockType = 1
rsGalleryAdd.Open()
MM_editRedirectUrl = "gallery_edit.asp?fldGalleryID=" & rsGalleryAdd.Fields.Item("fldGalleryID").Value
%>
<%
'Variables to Set
dim thisfilename
thisfilename = "gallery_add.asp"
dim filemanagerdir
filemanagerdir = "\images\" 'Relative to where the root of the website is
dim filemanagerdbdir
filemanagerdbdir = (rsGalleryAdd.Fields.Item("fldGalleryID").Value)
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
Else
	Set fs = CreateObject("Scripting.FileSystemObject") 
	Set a = fs.CreateFolder(Foldertocreate)  
	Set fs=nothing
End If

Response.Redirect(MM_editRedirectUrl)
%>