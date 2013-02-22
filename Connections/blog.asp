<%
' FileName="default_oledb.htm"
' Type="ADO" 
' DesigntimeType="ADO"
' HTTP="false"
' Catalog=""
' Schema=""
'Dim MM_blog_STRING
' Change your database connection string here:
MM_blog_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\wwwroot\tmp\Blog.mdb;Persist Security Info=False"
'MM_blog_STRING = "Driver={SQL Server};Server=Aron1;Database=pubs;Uid=sa;Pwd=asdasd;"

'************* No need to edit below here *******************

'Derive the path, no more manual config
'Dim theBasePath
theBasePath = Session("BlogPath") & "ckeditor/"  'FCKeditor base path
'Dim theConfigUserFilesPath
'theConfigUserFilesPath = "/dev/UserFiles/" 'If you have any problems, uncomment and set manually and comment out the next line
theConfigUserFilesPath = Session("BlogPath") & "UserFiles/"

if request("view") <> "" AND Session("MM_UserID") <> "" then
	if request("view") = 1 then
		Session("isAdmin") = 0
	elseif request("view") = 2 then 
		Session("isAdmin") = 1
	end if
end if

Function IsValidString(sValidate)
    Dim sInvalidChars
    Dim bTemp
    Dim i 
    ' Disallowed characters
    sInvalidChars = "!#$%^&*()=+{}[]|\\;?><'"
    for i = 1 To Len(sInvalidChars)
        if InStr(sValidate, Mid(sInvalidChars, i, 1)) > 0 then bTemp = True
        if bTemp then Exit For
    next
    for i = 1 to Len(sValidate)
        if Asc(Mid(sValidate, i, 1)) = 160 then bTemp = True
        if bTemp then Exit For
    next

    if not bTemp then
        bTemp = InStr(sValidate, "..") > 0
    end if
    if not bTemp then
        bTemp = InStr(sValidate, "  ") > 0
    end if
    if not bTemp then
        bTemp = (len(sValidate) <> len(Trim(sValidate)))
    end if 'Addition for leading and trailing spaces
    if not bTemp then
        bTemp = len(sValidate) < 1
    end if 'Empty      

    ' if any of the above are true, invalid string
    IsValidString = Not bTemp
End Function
function HackerSafe_Filter(cleanvar)
	        '  Encode Ampersand
	 cleanvar = replace(cleanvar,"&", "&")
	        '  Encode Single Quote
	 cleanvar = replace(cleanvar,"'", "'")
	        '  Encode Double Quote
	 'cleanvar = replace(cleanvar,""""", """)
	        '  Encode Less Than
	 cleanvar = replace(cleanvar,">", ">")
	        '  Encode Greater Than
	 cleanvar = replace(cleanvar,"<", "<")
	        '  Encode Close Bracket
	 cleanvar = replace(cleanvar,")", ")")
	        '  Encode Open Bracket
	 cleanvar = replace(cleanvar,"(", "(")
	        '  Encode Close Square Bracket
	 cleanvar = replace(cleanvar,"]", "]")
	        '  Encode Open Square Bracket
	 cleanvar = replace(cleanvar,"[", "[")
	        '  Encode Semicolon
	 cleanvar = replace(cleanvar,";", ";")
	        '  Encode Colon
	 cleanvar = replace(cleanvar,":", ":")
	        '  Encode Forward Slash
	 cleanvar = replace(cleanvar,"/", "/")
	        '  Encode Left Brace
	 cleanvar = replace(cleanvar,"}", "}")
	        '  Encode Right Brace
	 cleanvar = replace(cleanvar,"{", "{")
	        '  Encode Exclamation
	 cleanvar = replace(cleanvar,"!", "!")
	        '  Encode Double Dash
	 cleanvar = replace(cleanvar,"--", "--")
	        '  Encode Equal Sign
	 cleanvar = replace(cleanvar,"=", "=")
	        '  Encode Underscore
	 cleanvar = replace(cleanvar,"_", "_")
	 HackerSafe_Filter = cleanvar
end function
FUNCTION CropSentence(strText, intLength, strTrial) 
  Dim wsCount 
  Dim intTempSize 
  Dim intTotalLen 
  Dim strTemp 
  
  wsCount = 0 
  intTempSize = 0 
  intTotalLen = 0 
  intLength = intLength - Len(strTrial) 
  strTemp = "" 
    
  IF Len(strText) > intLength THEN 
    arrTemp = Split(strText, " ") 
    FOR EACH x IN arrTemp 
      IF Len(strTemp) <= intLength THEN 
        strTemp = strTemp & x & " " 
      END IF 
    NEXT 
      CropSentence = Left(strTemp, Len(strTemp) - 1) & strTrial 
  ELSE 
    CropSentence = strText 
  END IF 
END FUNCTION
function DoDateTime(str, nNamedFormat, nLCID)				
	dim strRet								
	dim nOldLCID								
										
	strRet = str								
	If (nLCID > -1) Then							
		oldLCID = Session.LCID						
	End If									
										
	On Error Resume Next							
										
	If (nLCID > -1) Then							
		Session.LCID = nLCID						
	End If									
										
	If ((nLCID < 0) Or (Session.LCID = nLCID)) Then				
		strRet = FormatDateTime(str, nNamedFormat)			
	End If									
										
	If (nLCID > -1) Then							
		Session.LCID = oldLCID						
	End If									
										
	DoDateTime = strRet							
End Function
function CI_StripHTML(strtext)				
 on error resume next	
 arysplit=split(strtext,"<")	
  if len(arysplit(0))>0 then j=1 else j=0	
  for i=j to ubound(arysplit)	
     if instr(arysplit(i),">") then	
       arysplit(i)=mid(arysplit(i),instr(arysplit(i),">")+1)	
     else	
       arysplit(i)="<" & arysplit(i)	
     end if	
  next	
  strOutput = join(arysplit, "")	
  strOutput = mid(strOutput, 2-j)	
  strOutput = replace(strOutput,">",">")	
  strOutput = replace(strOutput,"<","<")
  strOutput = replace(strOutput,"&quot;","")
  strOutput = replace(strOutput,"""","")
  strOutput = replace(strOutput,VbCrLf,"")
  strOutput = replace(strOutput,"&nbsp;","")
  CI_StripHTML = RemoveHTML(strOutput)
End Function

Function RemoveHTML(strText)
	Dim RegEx
	Set RegEx = New RegExp
	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True
	RemoveHTML = RegEx.Replace(strText, "")
End Function

Function MakeHyperlink(strSource)
'Thanks to Nate Rice http://www.naterice.com

  if (strSource <> "") or Not IsNull(strSource) then
    splitSource = Split(strSource," ")
    linkResult = Filter(splitSource, "www.", True)
    emailResult = Filter(splitSource, "@", True)

    For i = 0 to Ubound(splitSource)
	  'no clue what this does or why you'd want it...
	  'adding spaces back between the words. I think?
          'Not sure why you wouldn't just use "strSource"??
      MakeHyperLink = MakeHyperlink & splitSource(i) & " "
    Next

    For i = 0 to Ubound(linkResult)
      'http links...
      If Instr(linkResult(i), vbCRLF) > 0 Then
	'this is to check for splits of crlf instead of "space" as delimiter

        aMoreLinksSplit = Split(linkResult(i), vbCRLF)
        aMoreLinks = Filter(aMoreLinksSplit, "www.", True)

        For j = 0 to Ubound(aMoreLinks)
          'cutoff trailing periods or commas...
          If Right(aMoreLinks(j),1) = "," OR Right(aMoreLinks(j),1) = "." then
            aMoreLinks(j)=left(aMoreLinks(j),len(aMoreLinks(j))-1)
          End If
	
          'check for existance of http://
          If Left(lcase(aMoreLinks(j)),7) = "http://" then
            linkWrapTag = "<a href=""" & aMoreLinks(j) & """>" & aMoreLinks(j) & "</a>"
          Else
            linkWrapTag = "<a href=http://"""& aMoreLinks(j) & """>" & aMoreLinks(j) & "</a>"
          End If

          MakeHyperLink = Replace(MakeHyperLink, aMoreLinks(j), linkWrapTag)
        Next

      Else  
        'cutoff trailing periods or commas...
        If Right(linkResult(i),1) = "," OR Right(linkResult(i),1) = "." then
          linkResult(i)=left(linkResult(i),len(linkResult(i))-1)
        End If
	
        'check for existance of http://
        If Left(lcase(linkResult(i)),7) = "http://" then
          linkWrapTag = "<a href=""" & linkResult(i) & """>" & linkResult(i) & "</a>"
        Else
          linkWrapTag = "<a href=http://"""& linkResult(i) & """>" & linkResult(i) & "</a>"
        End If

        MakeHyperLink = Replace(MakeHyperLink, linkResult(i), linkWrapTag)
      End If
      
    Next

    For i = 0 to Ubound (emailResult)
      If Right(emailResult(i),1) = "," OR Right(emailResult(i),1) = "." then
        emailResult(i)=left(emailResult(i),len(emailResult(i))-1)
      End If
      emailWrapTag = "<a href=mailto:" & emailResult(i) & ">" & emailResult(i) & "</a>"

      MakeHyperLink = Replace(MakeHyperLink, emailResult(i), emailWrapTag)
    Next

    MakeHyperLink = Replace(MakeHyperLink,Chr(13),"<br>")

  Else
    MakeHyperlink = strSource
  End If
End Function
%>
<!--#include file="plugin.asp" -->