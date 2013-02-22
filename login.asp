<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
if IsValidString(Request.Form("username")) AND IsValidString(Request.form("remember")) AND IsValidString(Request.Form("password")) AND IsValidString(Request.QueryString("accessdenied")) Then
If cstr(Request.Form("username"))<>"" Then
  If Request.form("remember") ="1" Then
     Response.Cookies("ckUsername") = Request.Form("username")
     Response.Cookies("ckPassword") = Request.Form("password")
     Response.Cookies("ckRemember") = "1"
     Response.Cookies("ckUsername").expires = Date + 30
     Response.Cookies("ckPassword").expires = Date + 30
     Response.Cookies("ckRemember").expires = Date + 30
  Else
     Response.Cookies("ckRemember") = ""
     Response.Cookies("ckUsername") = ""
     Response.Cookies("ckPassword") = ""
  End If
End If
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="Approved"
  MM_redirectLoginSuccess="main.asp"
  MM_redirectLoginFailed="login.asp?lf=true"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_blog_STRING
  MM_rsUser.Source = "SELECT fldAuthorUsername, fldAuthorPassword"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM tblAuthor WHERE fldAuthorUsername='" & Replace(MM_valUsername,"'","''") &"' AND fldAuthorPassword='" & Replace(Request.Form("password"),"'","''") & "' AND Approved = 1"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
end if
%>
<html>
<head>
<title>Login</title>
<script type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function YY_checkform() { //v4.66
//copyright (c)1998,2002 Yaromat.com
  var args = YY_checkform.arguments; var myDot=true; var myV=''; var myErr='';var addErr=false;var myReq;
  for (var i=1; i<args.length;i=i+4){
    if (args[i+1].charAt(0)=='#'){myReq=true; args[i+1]=args[i+1].substring(1);}else{myReq=false}
    var myObj = MM_findObj(args[i].replace(/\[\d+\]/ig,""));
    myV=myObj.value;
    if (myObj.type=='text'||myObj.type=='password'||myObj.type=='hidden'){
      if (myReq&&myObj.value.length==0){addErr=true}
      if ((myV.length>0)&&(args[i+2]==1)){ //fromto
        var myMa=args[i+1].split('_');if(isNaN(myV)||myV<myMa[0]/1||myV > myMa[1]/1){addErr=true}
      } else if ((myV.length>0)&&(args[i+2]==2)){
          var rx=new RegExp("^[\\w\.=-]+@[\\w\\.-]+\\.[a-z]{2,4}$");if(!rx.test(myV))addErr=true;
      } else if ((myV.length>0)&&(args[i+2]==3)){ // date
        var myMa=args[i+1].split("#"); var myAt=myV.match(myMa[0]);
        if(myAt){
          var myD=(myAt[myMa[1]])?myAt[myMa[1]]:1; var myM=myAt[myMa[2]]-1; var myY=myAt[myMa[3]];
          var myDate=new Date(myY,myM,myD);
          if(myDate.getFullYear()!=myY||myDate.getDate()!=myD||myDate.getMonth()!=myM){addErr=true};
        }else{addErr=true}
      } else if ((myV.length>0)&&(args[i+2]==4)){ // time
        var myMa=args[i+1].split("#"); var myAt=myV.match(myMa[0]);if(!myAt){addErr=true}
      } else if (myV.length>0&&args[i+2]==5){ // check this 2
            var myObj1 = MM_findObj(args[i+1].replace(/\[\d+\]/ig,""));
            if(myObj1.length)myObj1=myObj1[args[i+1].replace(/(.*\[)|(\].*)/ig,"")];
            if(!myObj1.checked){addErr=true}
      } else if (myV.length>0&&args[i+2]==6){ // the same
            var myObj1 = MM_findObj(args[i+1]);
            if(myV!=myObj1.value){addErr=true}
      }
    } else
    if (!myObj.type&&myObj.length>0&&myObj[0].type=='radio'){
          var myTest = args[i].match(/(.*)\[(\d+)\].*/i);
          var myObj1=(myObj.length>1)?myObj[myTest[2]]:myObj;
      if (args[i+2]==1&&myObj1&&myObj1.checked&&MM_findObj(args[i+1]).value.length/1==0){addErr=true}
      if (args[i+2]==2){
        var myDot=false;
        for(var j=0;j<myObj.length;j++){myDot=myDot||myObj[j].checked}
        if(!myDot){myErr+='* ' +args[i+3]+'\n'}
      }
    } else if (myObj.type=='checkbox'){
      if(args[i+2]==1&&myObj.checked==false){addErr=true}
      if(args[i+2]==2&&myObj.checked&&MM_findObj(args[i+1]).value.length/1==0){addErr=true}
    } else if (myObj.type=='select-one'||myObj.type=='select-multiple'){
      if(args[i+2]==1&&myObj.selectedIndex/1==0){addErr=true}
    }else if (myObj.type=='textarea'){
      if(myV.length<args[i+1]){addErr=true}
    }
    if (addErr){myErr+='* '+args[i+3]+'\n'; addErr=false}
  }
  if (myErr!=''){alert('The required information is incomplete or contains errors:\t\t\t\t\t\n\n'+myErr)}
  document.MM_returnValue = (myErr=='');
}
//-->
</script>
</head>
<body>
 <h1 align="center"> Login</h1>
 <form action="<%=MM_LoginAction%>" method="POST" name="loginFrm" id="loginFrm">
<table width="500"  border="0" align="center" cellpadding="3" cellspacing="2" class="tabledisplay">
<% If (Request.QueryString("lf")) = ("true") Then 'script %>
<tr><th colspan="2">Login Failed, Try Again (or you may not be approved just yet)</th>
</tr>
<% End If ' end If (Request.QueryString("lf")) = ("true") script %>
<tr>
<th align="right">Username:</th>
<td>
<input value="<%= Request.Cookies("ckUsername") %>" name="username" type="text" class="txtBox" id="username" tabindex="1" /></td>
</tr>
<tr>
<th align="right">Password:</th>
<td><input value="<%= Request.Cookies("ckPassword") %>" name="password" type="password" class="txtBox" id="password" tabindex="2" /></td>
</tr>
<tr>
<th align="right">Remember Me:</th>
<td>
<input name="remember" type="checkbox" id="remember" tabindex="3" value="1" checked="checked" <%If (Request.Cookies("ckRemember") = "1") Then Response.Write("CHECKED") : Response.Write("")%> /></td>
</tr>
<tr align="center" valign="middle">
<td colspan="2">
  <input name="lgnBtn" type="submit" class="btn" id="lgnBtn" tabindex="4" onclick="YY_checkform('loginFrm','username','#q','0','Username Required','password','#q','0','Password Required');return document.MM_returnValue" value="Login" /></td>
</tr>
</table>
</form>
</body>
</html>