<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/blog.asp" -->
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


<%
Dim rsBlog
Dim rsBlog_numRows
Set rsBlog = Server.CreateObject("ADODB.Recordset")
rsBlog.ActiveConnection = MM_blog_STRING
rsBlog.Source = "SELECT * FROM tblPage ORDER BY PageName ASC"
rsBlog.CursorType = 0
rsBlog.CursorLocation = 2
rsBlog.LockType = 1
rsBlog.Open()
rsBlog_numRows = 0
%>
<%
Dim rs_blog2__MMColParam
rs_blog2__MMColParam = "4"
If (Request("MM_EmptyValue") <> "") Then 
  rs_blog2__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rs_blog2
Dim rs_blog2_numRows
Set rs_blog2 = Server.CreateObject("ADODB.Recordset")
rs_blog2.ActiveConnection = MM_blog_STRING
rs_blog2.Source = "SELECT * FROM tblBlog WHERE BlogCat = " + Replace(rs_blog2__MMColParam, "'", "''") + " ORDER BY BlogDate DESC"
rs_blog2.CursorType = 0
rs_blog2.CursorLocation = 2
rs_blog2.LockType = 1
rs_blog2.Open()
rs_blog2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 20
Repeat1__index = 0
rsBlog_numRows = rsBlog_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index
Repeat2__numRows = 20
Repeat2__index = 0
rs_blog2_numRows = rs_blog2_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables
' set the record count
rsBlog_total = rsBlog.RecordCount
' set the number of rows displayed on this page
If (rsBlog_numRows < 0) Then
  rsBlog_numRows = rsBlog_total
Elseif (rsBlog_numRows = 0) Then
  rsBlog_numRows = 1
End If
' set the first and last displayed record
rsBlog_first = 1
rsBlog_last  = rsBlog_first + rsBlog_numRows - 1
' if we have the correct record count, check the other stats
If (rsBlog_total <> -1) Then
  If (rsBlog_first > rsBlog_total) Then rsBlog_first = rsBlog_total
  If (rsBlog_last > rsBlog_total) Then rsBlog_last = rsBlog_total
  If (rsBlog_numRows > rsBlog_total) Then rsBlog_numRows = rsBlog_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them
If (rsBlog_total = -1) Then
  ' count the total records by iterating through the recordset
  rsBlog_total=0
  While (Not rsBlog.EOF)
    rsBlog_total = rsBlog_total + 1
    rsBlog.MoveNext
  Wend
  ' reset the cursor to the beginning
  If (rsBlog.CursorType > 0) Then
    rsBlog.MoveFirst
  Else
    rsBlog.Requery
  End If
  ' set the number of rows displayed on this page
  If (rsBlog_numRows < 0 Or rsBlog_numRows > rsBlog_total) Then
    rsBlog_numRows = rsBlog_total
  End If
  ' set the first and last displayed record
  rsBlog_first = 1
  rsBlog_last = rsBlog_first + rsBlog_numRows - 1
  If (rsBlog_first > rsBlog_total) Then rsBlog_first = rsBlog_total
  If (rsBlog_last > rsBlog_total) Then rsBlog_last = rsBlog_total
End If
%>
<%
' *** Move To Record and Go To Record: declare variables
Set MM_rs    = rsBlog
MM_rsCount   = rsBlog_total
MM_size      = rsBlog_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter
if (Not MM_paramIsDefined And MM_rsCount <> 0) then
  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)
  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If
  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record
End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range
If (MM_rsCount = -1) Then
  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend
  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If
  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If
  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If
  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats
' set the first and last displayed record
rsBlog_first = MM_offset + 1
rsBlog_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsBlog_first > MM_rsCount) Then rsBlog_first = MM_rsCount
  If (rsBlog_last > MM_rsCount) Then rsBlog_last = MM_rsCount
End If
' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters
' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""
' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next
' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next
' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links
MM_keepMove = MM_keepBoth
MM_moveParam = "index"
' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If
' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Pages</title>
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
		<li><a class="current" href="pages.asp">Pages</a></li>
			<ul id="secondary">
				<li><a href="add_page.asp">Create a New Page</a></li>
			</ul>
		<li><a href="cat.asp">Categories</a></li>
		<li><a href="users.asp">Users</a></li>
		<li><a href="layout.asp">Layout</a></li>
		<li><a href="blog_config.asp">Configuration</a></li>
		<% end if %>
	</ul>
	</div>
	<div id="main">
		<div id="contents">
      <h2>Pages</h2>
    <table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
      <tr>
        <th>Page Name (Title) </th>
        <th>Update / Delete</th>
      </tr>
      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsBlog.EOF)) 
%>
      <tr<% if Repeat1__index MOD 2 = 0 then response.Write(" class='alt'") end if%>>
        <td><a href="template.asp?pagename=<%=LCase(rsBlog.Fields.Item("PageName").Value)%>" title="View in your layout" target="_blank"><%=LCase(rsBlog.Fields.Item("PageName").Value)%> (<%=(rsBlog.Fields.Item("PageTitle").Value)%>)</a></td>
        <td><a href="update_page.asp?pagename=<%=(rsBlog.Fields.Item("PageName").Value)%>">Update</a><% if rsBlog.Fields.Item("PageName").Value  <> "thanks" then %> / <a href="delete_page.asp?pagename=<%=(rsBlog.Fields.Item("PageName").Value)%>">Delete</a><%end if%></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsBlog.MoveNext()
Wend
%>
      <tr>
        <td colspan="3">
          <%
For i = 1 to rsBlog_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsBlog_total Then TM_endCount = rsBlog_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsBlog_total) then Response.Write(", ")
next
 %>
         
      </td>
      </tr>
    </table>
	
    
		</div>
	</div>
</body>
</html>
<%
rsBlog.Close()
Set rsBlog = Nothing
%>
<%
rs_blog2.Close()
Set rs_blog2 = Nothing
%>

