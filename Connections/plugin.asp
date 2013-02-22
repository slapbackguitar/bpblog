<%
'Instructions: 
'1. Add the if/then statement referencing your plugin function name
'2. Add your function (remember to use page checking if needed)
function CheckPlugin(plugincode)
tmpText = plugincode
	if instr(tmpText, "%%plugin_name%%") > 0 then
		tmpText = Replace(tmpText, "%%plugin_name%%", plugin_name())
	end if
	if instr(tmpText, "%%plugin_userlastposts%%") > 0 then
		tmpText = Replace(tmpText, "%%plugin_userlastposts%%", plugin_userlastposts())
	end if		
	CheckPlugin = tmpText
end function
%>
<%
'--------------------------------------
public function plugin_name()
'Author:
'URL: 
'Description:
plugin_name = "<strong>Plugin Test</strong>"

end function
%>
<%
'-------------------------------------------------------------------------------------------
public function plugin_userlastposts()
'Author: Matt
'URL: http://blog.betaparticle.com
'Description: Diplays the user's last n posts on their profile page, n can be set below
'Template Code: %%plugin_userlastposts%%
userlastpostsnumber = 10 'Set number to display
	if Right(LCase(Request.ServerVariables("PATH_INFO")), 19) = "template_author.asp" then 'pagename where you want this to appear
		Dim tmpText
		Dim userlastposts
		Set userlastposts = Server.CreateObject("ADODB.Recordset")
		userlastposts.ActiveConnection = MM_blog_STRING
		userlastposts.Source = "SELECT TOP " & userlastpostsnumber & " * FROM tblBlog WHERE BlogAuthor = " + Request("id") + " AND BlogDraft = 0 ORDER BY BlogDate DESC"
		userlastposts.CursorType = 0
		userlastposts.CursorLocation = 2
		userlastposts.LockType = 1
		userlastposts.Open()
		if (NOT userlastposts.EOF) then
			tmpText = tmpText & "<h4>Last " & userlastpostsnumber & " Posts</h4>" & vbCrLf & "<ul>"
		end if
		While (NOT userlastposts.EOF) 
			tmpText = tmpText & "<li><a href='template_permalink.asp?id=" & userlastposts.Fields.Item("BlogID").Value & "'>" & userlastposts.Fields.Item("BlogHeadline").Value & "</a></li>"
			userlastposts.MoveNext()
		Wend
		if (NOT userlastposts.EOF) then
			tmpText = tmpText & vbCrLf & "</ul>"
		end if
		userlastposts.Close()
		Set userlastposts = Nothing
		plugin_userlastposts = tmpText
	else
		plugin_userlastposts = ""
	end if
end function

function readmore(BlogHTML, theid)

	BlogHTMLlen = len(BlogHTML)
	if instr(BlogHTML, "[more]") and theid > 0 then
		blnReadMore = true
		intbegin = instr(BlogHTML, "[more]")-1
		intend = instr(BlogHTML, "[more]")+6
	end if
	
	if blnReadMore then
		BlogHTMLtmp = mid(BlogHTML,1,(instr(BlogHTML, "[more]")-1))
		BlogHTMLtmp = BlogHTMLtmp & "<a href='template_permalink.asp?id=" & theid  &  "'>More...</a>"
		readmore = BlogHTMLtmp
	else
		if Right(LCase(Request.ServerVariables("PATH_INFO")), 22) = "template_permalink.asp" and instr(BlogHTML, "[more]") and theid = 0 then
			BlogHTMLtmp = replace(BlogHTML,"[more]","")
		else
			BlogHTMLtmp = BlogHTML
		end if
		readmore = BlogHTMLtmp
	end if

end function
%>