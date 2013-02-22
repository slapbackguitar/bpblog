<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/blog.asp" -->
<%
Dim rsLayoutConfig

Set rsLayoutConfig = Server.CreateObject("ADODB.Recordset")
rsLayoutConfig.ActiveConnection = MM_blog_STRING
rsLayoutConfig.Source = "SELECT blogLayout from tblBlogRSS"
rsLayoutConfig.CursorType = 0
rsLayoutConfig.CursorLocation = 2
rsLayoutConfig.LockType = 1
rsLayoutConfig.Open()

if request("layout") <> "" then
	If session("layout") <> "" Then
		layoutConfig = "Themes/" & session("layout") & "/search.asp"
	Else
		Response.End	
	End If
elseif session("layout") <> "" Then
	layoutConfig = "Themes/" & session("layout") & "/search.asp"	
else
	layoutConfig = "Themes/" & rsLayoutConfig.Fields.Item("blogLayout").Value & "/search.asp"
end if

Server.Execute(layoutConfig)
%>
<%
rsLayoutConfig.Close()
Set rsLayoutConfig = Nothing
%>