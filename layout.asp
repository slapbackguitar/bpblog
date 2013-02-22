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
Dim rsLayoutsConfig
Set rsLayoutsConfig = Server.CreateObject("ADODB.Recordset")
rsLayoutsConfig.ActiveConnection = MM_blog_STRING
rsLayoutsConfig.Source = "SELECT * FROM tblBlogRSS"
rsLayoutsConfig.CursorType = 0
rsLayoutsConfig.CursorLocation = 2
rsLayoutsConfig.LockType = 1
rsLayoutsConfig.Open()
%>

<%
Dim getFileName
getFileName = request.QueryString("file")
if getFileName = "" then getFileName = "default.asp"

Dim getAction
getAction = request.QueryString("action")
if getAction = "SaveFile" then
	getFileName = request("HiddenFileName")
	dim Str
	Str = request.form("CodeView")

	dim objFSO
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists (Server.MapPath(".") & "/Themes/" & rsLayoutsConfig.Fields.Item("blogLayout").value & "/" & getFileName) and Len(Str)>0 then
		dim objTextStream
		set objTextStream = objFSO.OpenTextFile ("" & Server.MapPath(".") & "/Themes/" & rsLayoutsConfig.Fields.Item("blogLayout").value & "/" & getFileName & "", 2, False, -1)
		objTextStream.write Str
		objTextStream.Close
		set objTextStream = Nothing
		Set objFSO = Nothing
	else
		response.write("Problem")
	end if
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Layout</title>
	<style type="text/css" media="screen">	@import "tabs.css";	</style>
    <link rel="stylesheet" href="css/validationEngine.jquery.css" type="text/css" media="screen" title="no title" charset="utf-8" />
    <script src="js/jquery.min.js" type="text/javascript"></script>
    <script src="js/jquery.validationEngine-en.js" type="text/javascript"></script>
    <script src="js/jquery.validationEngine.js" type="text/javascript"></script>
                                        <script>	
		$(document).ready(function() {		
			$("#saveit").validationEngine()
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
		<li><a href="users.asp">Users</a></li>
		<li><a class="current" href="layout.asp">Layout</a></li>
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">
          <h2>Layout</h2>
        <table width="99%" border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
        <colgroup>
        	<col width="300" />
            <col />
        </colgroup>
          <tr>
            <th>Files in theme</th>
            <th>Code</a></th>
          </tr>
          <tr style="vertical-align:top;">
          	<td>
<%
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set AFolder = FSO.GetFolder(Server.MapPath(".") & "/Themes/" & rsLayoutsConfig.Fields.Item("blogLayout").value)
For Each Item in AFolder.Files
  if getFileName = Item.name then
	 Response.Write "<a href=""layout.asp?file=" & Item.name & """ style=""background-color:#FFF3B3;"">" & Item.name & "</a><br />"
  else
  	 Response.Write "<a href=""layout.asp?file=" & Item.name & """>" & Item.name & "</a><br />"
  end if
Next
%>
            </td>
            <td>
            	<form method="post" action="?action=SaveFile" id="saveit">
<%
  dim objXMLHTTP
  
  Dim URL
  URL = Server.MapPath(".") & "/Themes/" & rsLayoutsConfig.Fields.Item("blogLayout").value & "/" & getFileName

  Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
  objXMLHTTP.Open "GET", URL, false
  objXMLHTTP.Send

  Response.Write "<h4>Code for "&getFileName&"</h4>"
  Response.Write "<textarea class=""validate[required]"" name=""CodeView"" id=""CodeView"" style=""width:90%;"" rows=""20"">"
  Response.Write objXMLHTTP.responseText
  Response.Write "</textarea>"
  Set objXMLHTTP = Nothing
%>
                <br /><br />
                <input type="hidden" name="HiddenFileName" value="<%=getFileName%>" />
                <input type="submit" value=" Update file " />
                <input type="reset" value=" Undo changes " />
                </form>
            </td>
          </tr>
        </table>
		</div>
	</div>
</body>
</html>
<%
rsLayoutsConfig.Close()
Set rsLayoutsConfig = Nothing
%>