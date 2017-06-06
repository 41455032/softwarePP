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
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_chris_STRING
    MM_editCmd.CommandText = "UPDATE dbo.login SET peo = ? WHERE peo = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("username"), Request.Form("username"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_chris_STRING
Recordset1_cmd.CommandText = "SELECT Peo_no, Peo_password FROM dbo.Peo" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_chris_STRING
Recordset2_cmd.CommandText = "SELECT * FROM dbo.login" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "manager.asp"
  MM_redirectLoginFailed = "Register.asp"

  MM_loginSQL = "SELECT Peo_no, Peo_password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.Peo WHERE Peo_no = ? AND Peo_password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_chris_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 5, 1, -1, MM_valUsername) ' adDouble
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 5, 1, -1, Request.Form("password")) ' adDouble
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
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
                    <h2>管理员请登录</h2>
                     <p style="text-align:left"><a href="Login.asp">如果你是学生，戳我</a></button></p>
                     <p style="text-align:left"><a href="LoginforT.asp">如果你是教师，戳我</a></button></p>
                </div>
               <form name="form1" action="<%=MM_editAction%>" method="POST">
                <div class="col-lg-10">
                    <input name="username" type="text" autofocus required class="form-control" placeholder="请输入账户名"/>
                </div>
                <div class="col-lg-10"></div>
                <div class="col-lg-10">
                    <input name="password" type="password" autofocus required class="form-control" placeholder="请输入密码"/>
                </div>
                <div class="col-lg-10"></div>
                <input name="submit" type="submit" value="登录" class="btn btn-success col-lg-8">
                <button class="btn-group col-lg-3"><a href="Register.asp">注册</a></button>
                <input type="hidden" name="MM_update" value="form1">
                <input type="hidden" name="MM_recordId" value="<%= Recordset2.Fields.Item("peo").Value %>">
               </form>
                
                <div class="col-lg-10"></div>
              
            </div>
        </div>
        

        </body>/
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
