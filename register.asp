<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/blog.asp" -->
<%
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

    ' if any of the above are true, invalid string
    IsValidString = Not bTemp
End Function

' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "loginFrm") AND isValidString(Request.Form("fldAuthorUsername")) = True AND isValidString(Request.Form("MM_insert")) = True Then

	Dim rsCheck
	Dim rsCheck_numRows
	
	Set rsCheck = Server.CreateObject("ADODB.Recordset")
	rsCheck.ActiveConnection = MM_blog_STRING
	rsCheck.Source = "SELECT * FROM tblAuthor WHERE fldAuthorUsername = '" + Request.Form("fldAuthorUsername") + "'"
	rsCheck.CursorType = 0
	rsCheck.CursorLocation = 2
	rsCheck.LockType = 1
	rsCheck.Open()
	
	rsCheck_numRows = 0
	Dim UsernameTaken
	UsernameTaken = 0
	if rsCheck.EOF then
		UsernameTaken = 0
	else
		UsernameTaken = 1
	end if

end if	
%>
<%
if UsernameTaken = 0 then
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

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>

<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

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

    ' if any of the above are true, invalid string
    IsValidString = Not bTemp
End Function


uname = trim(request.form("fldAuthorUsername")) 
pword =  trim(request.form("fldAuthorPassword")) 
email =  trim(request.form("fldAuthorEmail")) 
if isValidString(uname) = True AND isValidString(pword) = True AND isValidString(email) = True then

	If (CStr(Request("MM_insert")) <> "") Then
	
	
	  MM_editTable = "tblAuthor"
	  MM_editRedirectUrl = "template.asp?pagename=thankyou"
	  MM_fieldsStr  = "fldAuthorUsername|value|fldAuthorEmail|value|fldAuthorPassword|value"
	  MM_columnsStr = "fldAuthorUsername|',none,''|fldAuthorEmail|',none,''|fldAuthorPassword|',none,''"
	
	
	  MM_editQuery = "insert into tblAuthor (fldAuthorUsername, fldAuthorEmail, fldAuthorPassword) values ('" & uname & "', '" & pword & "', '" & email & "')"
	
		  If (Not MM_abortEdit) Then
			' execute the insert
			Set MM_editCmd = Server.CreateObject("ADODB.Command")
			MM_editCmd.ActiveConnection = MM_blog_STRING
			MM_editCmd.CommandText = MM_editQuery
				'Captcha start
				Function CheckCAPTCHA(valCAPTCHA)
					SessionCAPTCHA = Trim(Session("CAPTCHA"))
					Session("CAPTCHA") = vbNullString
					if Len(SessionCAPTCHA) < 1 then
						CheckCAPTCHA = False
						exit function
					end if
					if CStr(SessionCAPTCHA) = CStr(valCAPTCHA) then
						CheckCAPTCHA = True
					else
						CheckCAPTCHA = False
					end if
				End Function
				
				strCAPTCHA = Trim(Request.Form("strCAPTCHA"))
				
				if CheckCAPTCHA(strCAPTCHA) = false then
					response.Write("You did not type in the verification info correctly")
					Response.End
				end if
				'Captcha end    
				MM_editCmd.Execute
		
			MM_editCmd.ActiveConnection.Close
		
			If (MM_editRedirectUrl <> "") Then
			  Response.Redirect(MM_editRedirectUrl)
			End If
		  End If
		
		End If
	end if
else
	response.Write("Something was wrong with your submission.  Either your username, password or email wasn't in the correct format.")
	Response.End
end if	
%>
<%
Dim rsBlogSite
Dim rsBlogSite_numRows

Set rsBlogSite = Server.CreateObject("ADODB.Recordset")
rsBlogSite.ActiveConnection = MM_blog_STRING
rsBlogSite.Source = "SELECT * FROM tblBlogRSS"
rsBlogSite.CursorType = 0
rsBlogSite.CursorLocation = 2
rsBlogSite.LockType = 1
rsBlogSite.Open()

rsBlogSite_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat1__numRows
%>

							
<script language="VBScript" type="text/vbscript" runat="server">										
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
  CI_StripHTML = strOutput	
End Function														
</script>
<title><%=(rsBlogSite.Fields.Item("blogTitle").Value)%> | Register</title>
<meta name="Description" content="Archives by Month" />
<link rel="alternate" type="application/rss+xml" href="rss.xml" title="RSS feed for <%=(rsBlogSite.Fields.Item("blogTitle").Value)%>">
<script type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function YY_checkform() { //v4.71
//copyright (c)1998,2002 Yaromat.com
  var a=YY_checkform.arguments,oo=true,v='',s='',err=false,r,o,at,o1,t,i,j,ma,rx,cd,cm,cy,dte,at;
  for (i=1; i<a.length;i=i+4){
    if (a[i+1].charAt(0)=='#'){r=true; a[i+1]=a[i+1].substring(1);}else{r=false}
    o=MM_findObj(a[i].replace(/\[\d+\]/ig,""));
    o1=MM_findObj(a[i+1].replace(/\[\d+\]/ig,""));
    v=o.value;t=a[i+2];
    if (o.type=='text'||o.type=='password'||o.type=='hidden'){
      if (r&&v.length==0){err=true}
      if (v.length>0)
      if (t==1){ //fromto
        ma=a[i+1].split('_');if(isNaN(v)||v<ma[0]/1||v > ma[1]/1){err=true}
      } else if (t==2){
        rx=new RegExp("^[\\w\.=-]+@[\\w\\.-]+\\.[a-zA-Z]{2,4}$");if(!rx.test(v))err=true;
      } else if (t==3){ // date
        ma=a[i+1].split("#");at=v.match(ma[0]);
        if(at){
          cd=(at[ma[1]])?at[ma[1]]:1;cm=at[ma[2]]-1;cy=at[ma[3]];
          dte=new Date(cy,cm,cd);
          if(dte.getFullYear()!=cy||dte.getDate()!=cd||dte.getMonth()!=cm){err=true};
        }else{err=true}
      } else if (t==4){ // time
        ma=a[i+1].split("#");at=v.match(ma[0]);if(!at){err=true}
      } else if (t==5){ // check this 2
            if(o1.length)o1=o1[a[i+1].replace(/(.*\[)|(\].*)/ig,"")];
            if(!o1.checked){err=true}
      } else if (t==6){ // the same
            if(v!=MM_findObj(a[i+1]).value){err=true}
      }
    } else
    if (!o.type&&o.length>0&&o[0].type=='radio'){
          at = a[i].match(/(.*)\[(\d+)\].*/i);
          o2=(o.length>1)?o[at[2]]:o;
      if (t==1&&o2&&o2.checked&&o1&&o1.value.length/1==0){err=true}
      if (t==2){
        oo=false;
        for(j=0;j<o.length;j++){oo=oo||o[j].checked}
        if(!oo){s+='* '+a[i+3]+'\n'}
      }
    } else if (o.type=='checkbox'){
      if((t==1&&o.checked==false)||(t==2&&o.checked&&o1&&o1.value.length/1==0)){err=true}
    } else if (o.type=='select-one'||o.type=='select-multiple'){
      if(t==1&&o.selectedIndex/1==0){err=true}
    }else if (o.type=='textarea'){
      if(v.length<a[i+1]){err=true}
    }
    if (err){s+='* '+a[i+3]+'\n'; err=false}
  }
  if (s!=''){alert('The required information is incomplete or contains errors:\t\t\t\t\t\n\n'+s)}
  document.MM_returnValue = (s=='');
}
//-->
</script>

<h1 id="blog-title"><a href="default.asp" accesskey="1"><%=(rsBlogSite.Fields.Item("blogTitle").Value)%></a></h1>
<p id="description"><%=(rsBlogSite.Fields.Item("blogSubTitle").Value)%></p>

<h3>Register</h3>
<form action="<%=MM_editAction%>" method="POST" name="loginFrm" id="loginFrm">
  <table width="100%"  border="0" align="center" cellpadding="3" cellspacing="2" class="tabledisplay">
    <% If UsernameTaken = 1 Then %>
    <tr>
      <th colspan="2">The username &quot;<%=request("fldAuthorUsername")%>&quot; is already taken, please try again</th>
    </tr>
    <% End If  %>
    <tr>
      <th align="right">Desired Username:</th>
      <td><input name="fldAuthorUsername" type="text" class="txtBox" id="fldAuthorUsername" tabindex="1" value="" /></td>
    </tr>
    <tr>
      <th align="right">Email Address: </th>
      <td><input name="fldAuthorEmail" type="text" class="txtBox" id="fldAuthorEmail" tabindex="1" value="<%=request("fldAuthorEmail")%>" /></td>
    </tr>
    <tr>
      <th align="right">Password:</th>
      <td><input name="fldAuthorPassword" type="password" class="txtBox" id="fldAuthorPassword" tabindex="2" value="<%=request("fldAuthorPassword")%>" /></td>
    </tr>
    <tr>
      <th align="right">Password Again:</th>
      <td><input name="fldAuthorPassword2" type="password" class="txtBox" id="fldAuthorPassword2" tabindex="2" value="<%=request("fldAuthorPassword")%>" /></td>
    </tr>
    <tr align="center" valign="middle">
      <td colspan="2"><p><img src='aspcaptcha.asp' alt='CAPTCHA' width='86' height='21' /><br />Write the characters in the image above<br /><input name='strCAPTCHA' type='text' id='strCAPTCHA' maxlength='8' /></p></td>
    </tr>
    <tr align="center" valign="middle">
      <td colspan="2"><input name="lgnBtn" type="submit" class="btn" id="lgnBtn" tabindex="4" value="Register &gt;&gt;" onclick="YY_checkform('loginFrm','fldAuthorUsername','#q','0','Username Required','fldAuthorEmail','S','2','Email must be valid','strCAPTCHA','#q','0','CAPTCHA Required','fldAuthorPassword','#fldAuthorPassword2','6','Passwords must match!');return document.MM_returnValue" /></td>
    </tr>
  </table>

  <input type="hidden" name="MM_insert" value="loginFrm">
</form>


<%
rsBlogSite.Close()
Set rsBlogSite = Nothing
%>
