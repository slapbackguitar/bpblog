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
Dim rs_cat
Dim rs_cat_numRows
Set rs_cat = Server.CreateObject("ADODB.Recordset")
rs_cat.ActiveConnection = MM_blog_STRING
rs_cat.Source = "SELECT * FROM tblCat ORDER BY CatName DESC"
rs_cat.CursorType = 0
rs_cat.CursorLocation = 2
rs_cat.LockType = 1
rs_cat.Open()
rs_cat_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = -1
Repeat1__index = 0
rs_cat_numRows = rs_cat_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables
Dim rs_cat_total
Dim rs_cat_first
Dim rs_cat_last
' set the record count
rs_cat_total = rs_cat.RecordCount
' set the number of rows displayed on this page
If (rs_cat_numRows < 0) Then
  rs_cat_numRows = rs_cat_total
Elseif (rs_cat_numRows = 0) Then
  rs_cat_numRows = 1
End If
' set the first and last displayed record
rs_cat_first = 1
rs_cat_last  = rs_cat_first + rs_cat_numRows - 1
' if we have the correct record count, check the other stats
If (rs_cat_total <> -1) Then
  If (rs_cat_first > rs_cat_total) Then
    rs_cat_first = rs_cat_total
  End If
  If (rs_cat_last > rs_cat_total) Then
    rs_cat_last = rs_cat_total
  End If
  If (rs_cat_numRows > rs_cat_total) Then
    rs_cat_numRows = rs_cat_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them
If (rs_cat_total = -1) Then
  ' count the total records by iterating through the recordset
  rs_cat_total=0
  While (Not rs_cat.EOF)
    rs_cat_total = rs_cat_total + 1
    rs_cat.MoveNext
  Wend
  ' reset the cursor to the beginning
  If (rs_cat.CursorType > 0) Then
    rs_cat.MoveFirst
  Else
    rs_cat.Requery
  End If
  ' set the number of rows displayed on this page
  If (rs_cat_numRows < 0 Or rs_cat_numRows > rs_cat_total) Then
    rs_cat_numRows = rs_cat_total
  End If
  ' set the first and last displayed record
  rs_cat_first = 1
  rs_cat_last = rs_cat_first + rs_cat_numRows - 1
  If (rs_cat_first > rs_cat_total) Then
    rs_cat_first = rs_cat_total
  End If
  If (rs_cat_last > rs_cat_total) Then
    rs_cat_last = rs_cat_total
  End If
End If
%>
<%
Dim MM_paramName
%>
<%
' *** Move To Record and Go To Record: declare variables
Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined
Dim MM_param
Dim MM_index
Set MM_rs    = rs_cat
MM_rsCount   = rs_cat_total
MM_size      = rs_cat_numRows
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
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If
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
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If
End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range
If (MM_rsCount = -1) Then
  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
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
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats
' set the first and last displayed record
rs_cat_first = MM_offset + 1
rs_cat_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rs_cat_first > MM_rsCount) Then
    rs_cat_first = MM_rsCount
  End If
  If (rs_cat_last > MM_rsCount) Then
    rs_cat_last = MM_rsCount
  End If
End If
' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters
Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth
Dim MM_removeList
Dim MM_item
Dim MM_nextItem
' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If
MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""
' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next
' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next
' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If
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
Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev
Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam
MM_keepMove = MM_keepBoth
MM_moveParam = "index"
' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If
' set the strings for the move to links
If (MM_keepMove <> "") Then
  MM_keepMove = MM_keepMove & "&"
End If
MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<title>Categories</title>

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
      <h2>Categories</h2>
      <table border="0" cellpadding="0" cellspacing="1" class="tabledisplay">
        <tr>
          <th>Category</th>
          <th>Description</th>
          <th>Options</th>
        </tr>
 <%
While ((Repeat1__numRows <> 0) AND (NOT rs_cat.EOF))
%>
        <tr<% if Repeat1__index MOD 2 = 0 then response.Write(" class='alt'") end if%>>
          <td align="default"> <%=(rs_cat.Fields.Item("CatName").Value)%> </td>
          <td align="default"><%=(rs_cat.Fields.Item("CatDesc").Value)%></td>
          <td align="default"><a href="cat_update.asp?catID=<%=(rs_cat.Fields.Item("CatID").Value)%>">Edit</a>/<a href="cat_delete.asp?catID=<%=(rs_cat.Fields.Item("CatID").Value)%>">Delete</a></td>
        </tr>
        <%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_cat.MoveNext()
Wend
%>
      </table>
      
		</div>
	</div>
</body>
</html>
<%
rs_cat.Close()
Set rs_cat = Nothing
%>

