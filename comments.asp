<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/blog.asp" -->

<%
If Request.form("remember") ="1" Then
Response.Cookies("txtName") = Request("txtName")
Response.Cookies("txtUrl") = Request("txtURL")
Response.Cookies("txtEmail") = Request("txtEmail")
Response.Cookies("remember") = "1"
Response.Cookies("txtName").Expires = Date + 365
Response.Cookies("txtURL").Expires = Date + 365
Response.Cookies("txtEmail").Expires = Date + 365
Response.Cookies("remember").expires = Date + 365
Else
Response.Cookies("txtName") = ""
Response.Cookies("txtURL") = ""
Response.Cookies("txtEmail") = ""
Response.Cookies("txtRemember") = ""
End If
%>

<%
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

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
MM_editRedirectUrl = "template.asp?pagename=thanks"
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_blog_STRING
  MM_editTable = "tblComment"
  'MM_editRedirectUrl = "thanks.html"
  MM_fieldsStr  = "txtName|value|txtURL|value|txtEmail|value|textarea|value|hiddenField|value"
  MM_columnsStr = "commentName|',none,''|commentURL|',none,''|commentEmail|',none,''|commentHTML|',none,''|blogID|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>

<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
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

      Response.Redirect(MM_editRedirectUrl)


  End If

End If
%>

<%
Dim rsComments__MMColParam
rsComments__MMColParam = "0"
%>

