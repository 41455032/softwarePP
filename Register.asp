<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/chris.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Redirect if username exists
MM_flag = "MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  Dim MM_rsKey
  Dim MM_rsKey_cmd
  
  MM_dupKeyRedirect = "Login.asp"
  MM_dupKeyUsernameValue = CStr(Request.Form("username"))
  Set MM_rsKey_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsKey_cmd.ActiveConnection = MM_chris_STRING
  MM_rsKey_cmd.CommandText = "SELECT Peo_no FROM dbo.Peo WHERE Peo_no = ?"
  MM_rsKey_cmd.Prepared = true
  MM_rsKey_cmd.Parameters.Append MM_rsKey_cmd.CreateParameter("param1", 5, 1, -1, MM_dupKeyUsernameValue) ' adDouble
  Set MM_rsKey = MM_rsKey_cmd.Execute
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then 
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1, MM_dupKeyRedirect, "?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_chris_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Peo (Peo_no, Peo_password, Peo_phone) VALUES (?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("username"), Request.Form("username"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("password"), Request.Form("password"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("phone"), Request.Form("phone"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "TheFirstPage1.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_chris_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Peo" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Login</title>

<!-- Bootstrap -->
<link href="CSS/bootstrap.css" rel="stylesheet">
<script src="js/jquery.min.js"></script>
<script src="js/bootstrap.js"></script>
<link href="CSS/login style.css" rel="stylesheet" type="text/css">
<!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
<!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
<!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
</head>
<body>

            <div class="mycenter">
            <div class="mysign">
                <div class="col-lg-11 text-center text-info">
                    <h3>请注册您的账号</h3>
                </div>
              <form name="form1" action="<%=MM_editAction%>" method="POST">
                <div class="col-lg-10">
                    <input name="username" type="text" autofocus required class="form-control" placeholder="请输入学号"/>
                </div>
                <div class="col-lg-10"></div>
                <div class="col-lg-10">
                    <input name="password" type="password" autofocus required class="form-control" placeholder="请输入密码"/>
                </div>
                <div class="col-lg-10"></div>
                
              <div class="col-lg-10">
                    <input name="phone" type="text" autofocus required class="form-control" placeholder="请输入手机号"/>
                </div>
                <div class="col-lg-10"></div>
              
               
                <div class="col-lg-10"></div>
                <div class="col-lg-10">
                    <input name="submitt" type="submit" value="注册">
                </div>
                <input type="hidden" name="MM_insert" value="form1">
              </form>
            </div>
        </div>

</body>/
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
