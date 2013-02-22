<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/blog.asp" -->
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
If IsValidString(request("search")) = True Then
	searchterm = HackerSafe_Filter(Request("search"))
Else
	Response.End	
End If
%>	
<%
Dim rsSearch
Dim rsSearch_numRows

Set rsSearch = Server.CreateObject("ADODB.Recordset")
rsSearch.ActiveConnection = MM_blog_STRING
rsSearch.Source = "SELECT BlogID, BlogHeadline, BlogHTML, BlogDate, BlogCat, BlogAuthor, BlogCommentInclude, BlogReadMore, BlogDraft, CatID, CatName, CatDesc, fldAuthorID, fldAuthorRealName, (SELECT COUNT(*) FROM tblComment WHERE tblComment.BlogID = tblBlog.BlogID AND tblComment.CommentInclude = 1) as CommentCount, (SELECT COUNT(*) FROM tblBlog WHERE BlogCat = CatID) as CategoryCount   FROM tblBlog, tblCat, tblAuthor   WHERE (BlogCat = CatID) AND (BlogHeadline LIKE '%" + searchterm + "%' OR BlogHTML LIKE '% " + searchterm + "%') AND (tblBlog.BlogAuthor = tblAuthor.fldAuthorID) AND tblBlog.BlogDraft <> 1 ORDER BY BlogDate DESC"
rsSearch.CursorType = 0
rsSearch.CursorLocation = 2
rsSearch.LockType = 1
rsSearch.Open()

rsSearch_numRows = 0
%>
<%

Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsBlogSite_numRows = rsBlogSite_numRows + Repeat1__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 20
Repeat3__index = 0
rsSearch_numRows = rsSearch_numRows + Repeat3__numRows
%>
<%
' Dim the variables used in these 3 subroutines...
Dim objKeywordRegExp, i, j, kwd, matchPos1, matchPos2, matchPos3, badChar, ptrnBegin, ptrnEnd, tempStoreKeyword, intKeywordMatches
Set objKeywordRegExp = New RegExp
	objKeywordRegExp.IgnoreCase = True
	objKeywordRegExp.Global = True

Function highlight(StringToHighlight, kwdStringOrArray)

	' ***** Set variables to initial states *****
	'                                           *
	kwd = 0
	matchPos1 = 0
	matchPos2 = 0
	ptrnBegin = "()("
	ptrnEnd = ")()"
	intKeywordMatches = 0
	'                                           *
	' *******************************************

	' ***** Highlight one string *****
	'                                *
	If Not IsArray(kwdStringOrArray) Then
	' In this case, the whole keyword string is treated as one big keyword.
		StringToHighlight = RegExpReplace(StringToHighlight, Replace(kwdStringOrArray, "&quot;", """"))
	'                                *
	' ********************************
	
	Else
		' ***** Highlight all of the keywords in an array *****
		'                                                     *
		' ----- Find first matching keyword -----
		'                                       |
		' First, arrange the keywords from longest to shortest.
		Call SortLengths(kwdStringOrArray)
		
		' Then, find the first match.
		Do While (kwd <= UBound(kwdStringOrArray)) And (InStr(StringToHighlight, "<span") = 0)
			StringToHighlight = RegExpReplace(StringToHighlight, kwdStringOrArray(kwd))
			kwd = kwd + 1
		Loop
		'                                       |
		' --------------------------------------- 
		
		' ----- Cycle through remaining keywords ----- 
		'                                            |
		For i = kwd to UBound(kwdStringOrArray)
			' Reset the pattern bits
			ptrnBegin = "()("
			ptrnEnd = ")()"
			
			' See if the current keyword exists in the span.
			matchPos1 = InStr("<span class='hl'></span>", LCase(kwdStringOrArray(i)))
			matchPos2 = matchPos1 + Len(kwdStringOrArray(i))
			matchPos1 = matchPos1 - 1
			' Find the keyword in the string from the other end too.
			matchPos3 = InstrRev("<span class='hl'></span>", LCase(kwdStringOrArray(i))) - 1
			' If the keyword is in the span text more than once, matchPos3 and matchPos1 will not be the same.
			
			if matchPos1 > -1 then
				' The keyword is in the span text...

				if matchPos3 = matchPos1 then
					' ----- The keyword is in the span text only once -----
					'                                                     |
					' If the character before the keyword is not a space, add it to the beginning of the RegExp pattern string
					' to create a version of the keyword that should be excluded from highlighting, such as "<span".
					badChar = Mid("<span class='hl'></span>", matchPos1, 1)
					if badChar <> " " then ptrnBegin = "([^" & badChar & "])("
						
					' If the character after the keyword is not a space, add it to the end of the RegExp pattern string
					' to create a version of the keyword that should be excluded from highlighting, such as "class=".
					badChar = Mid("<span class='hl'></span>", matchPos2, 1)
					if badChar <> " " then ptrnEnd = ")([^" & badChar & "])"
						
					' Put the pattern together and highlight matches.
					StringToHighlight = RegExpReplace(StringToHighlight, kwdStringOrArray(i))
					'                                                     |
					' -----------------------------------------------------
				
				end if
				' If the keyword is in the span text more than once,
				' do not attempt to highlight the matched word in the text. Do nothing instead.
				
			else
				' ----- The keyword is not in the span text -----
				'                                               |
				' It is safe to highlight the matching words.				
				StringToHighlight = RegExpReplace(StringToHighlight, kwdStringOrArray(i))
				'                                               |
				' -----------------------------------------------
			end if
		
		Next 'i
	'                                                     *
	' *****************************************************
	End If
	
	highlight = StringToHighlight
End Function ' highlight

' ***** This RegExp escapes characters that have special meaning within RegExp patterns
Dim objLiteralsRegExp
	Set objLiteralsRegExp = New RegExp
	objLiteralsRegExp.Global = True
	objLiteralsRegExp.Pattern = "(\+|\?|\$|\(|\)|\.|\||\{|\}|\[|\]|\^|\*|\\)"

' ***** This is the function that actually highlights the matched words
' ***** in the string by replacing them with themselves wrapped in a span tag.
Function RegExpReplace(strHighlight, strKwd)
	' Piece the RegExp pattern together.
	objKeywordRegExp.Pattern = ptrnBegin & objLiteralsRegExp.Replace(strKwd, "\$1") & ptrnEnd
	' Highlight the matched keyword.
	RegExpReplace = objKeywordRegExp.Replace(strHighlight, "$1<span class='hl'>$2</span>$3")
	' If something was highlighted, increment the total number of keywords highlighted.
	if RegExpReplace <> strHighlight then intKeywordMatches = intKeywordMatches + 1
End Function

' ***** This subroutine uses a bubble sort to arrange
' ***** the keywords in the array by length
Sub SortLengths(aryKwds)
	For i = 0 To UBound(aryKwds) - 1
		For j = i + 1 To UBound(aryKwds)
			if Len(aryKwds(i)) < Len(aryKwds(j)) Then
				tempStoreKeyword = aryKwds(i)
				aryKwds(i) = aryKwds(j)
				aryKwds(j) = tempStoreKeyword
			End if 'i > j
		Next 'j
	Next 'j
End Sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> |  Search results for "<%=searchterm%>"</title>
<meta name="Description" content="<%=(rsBlogSite.Fields.Item("blogDesc").Value)%>" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="js/lightbox.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<link href="Themes/White Minima/styles-site.css" rel="stylesheet" type="text/css" />
</head>
<BODY>
<DIV id=container>
<DIV id=header>
<h1 id="blog-title"><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a></h1>
<p id="description"><%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></p>
</DIV>
<DIV id=content>
<DIV id=main>
<% if rsSearch.EOF then %>  
<div class="post">
  <div class="post-body">
  <p>Your search term &quot;<%=searchterm%>&quot; returned no results.  Please try another word or phrase.</p>
  </div>
</div>
<% Else %>
  <% 
While ((Repeat3__numRows <> 0) AND (NOT rsSearch.EOF)) 
%>
<% if lastid <> rsSearch.Fields.Item("BlogID").Value then %>
<% if lastdate <> DoDateTime((rsSearch.Fields.Item("BlogDate").Value), 1, 1033) then%>
<h2 class="date-header"><%= DoDateTime((rsSearch.Fields.Item("BlogDate").Value), 1, 1033) %></h2>
<%end if%>
<!-- Begin post -->
<div class="post"> 
  <h3 class="post-title" id="<%=(rsSearch.Fields.Item("BlogID").Value)%>"><a href="template_permalink.asp?id=<%=(rsSearch.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsSearch.Fields.Item("BlogHeadline").Value)%>"><%=highlight((rsSearch.Fields.Item("BlogHeadline").Value),searchterm)%></a></h3>
  <div class="post-body">
  <p><% =highlight(CropSentence(CI_StripHTML(rsSearch.Fields.Item("BlogHTML").Value), 250, "..."),searchterm) %></p>
  </div>
<p class="post-footer">Posted by <a href="template_author.asp?id=<%=(rsSearch.Fields.Item("fldAuthorID").Value)%>" title="<%=(rsSearch.Fields.Item("fldAuthorRealName").Value)%>'s Profile"><%=(rsSearch.Fields.Item("fldAuthorRealName").Value)%></a> at <a href="template_permalink.asp?id=<%=(rsSearch.Fields.Item("BlogID").Value)%>" title="Permalink for <%=(rsSearch.Fields.Item("BlogHeadline").Value)%>"><%= DoDateTime((rsSearch.Fields.Item("BlogDate").Value), 3, 1033) %></a> in <a href="template_archives_cat.asp?cat=<%=(rsSearch.Fields.Item("CatID").Value)%>" title="<%=(rsSearch.Fields.Item("CatDesc").Value)%>"><%=(rsSearch.Fields.Item("CatName").Value)%> (<%=(rsSearch.Fields.Item("CategoryCount").Value)%>)</a> | <a href="template_permalink.asp?id=<%=(rsSearch.Fields.Item("BlogID").Value)%>#comments">Comments (<%=(rsSearch.Fields.Item("CommentCount").Value)%>)</a>  </p>
</div><!-- End post -->  
  <% end if %>
  <% Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  lastdate = DoDateTime((rsSearch.Fields.Item("BlogDate").Value), 1, 1033)
  lastid = rsSearch.Fields.Item("BlogID").Value
  rsSearch.MoveNext()
Wend
%>
<% end if %>
</DIV>
<DIV id=sidebar>
<!--#include file="../../inc_sidebar.asp" -->
</DIV>
 </DIV>
<DIV id=footer><P>Powered by <a href="http://blog.betaparticle.com" title="Powered by BP Blog 9.0">BP Blog 9.0</a>
		 | <a href="rss.xml">Feed (RSS)</a></P>
</DIV>	
</DIV>	 
</BODY>
</HTML>
<%
rsBlogSite.Close()
Set rsBlogSite = Nothing
%><%
rsSearch.Close()
Set rsSearch = Nothing
%>