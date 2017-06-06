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
    MM_editCmd.CommandText = "UPDATE dbo.Peo SET Peo_no = ?, Peo_password = ?, Peo_name = ?, Peo_phone = ? WHERE Peo_no = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("username"), Request.Form("username"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("password"), Request.Form("password"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, -1, Request.Form("name")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("phone"), Request.Form("phone"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Peo WHERE Peo_no in(SELECT * FROM dbo.login)" 
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
<title>修改个人信息</title>

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
                    <h3>你可以修改个人信息</h3>
                </div>
              <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                <div class="col-lg-10">
               编号： <input name="username" type="text" autofocus required class="form-control" value="<%=(Recordset1.Fields.Item("Peo_no").Value)%>" readonly>
                   
                </div>
                <div class="col-lg-10"></div>
                <div class="col-lg-10">
                   密码： <input name="password" type="text" autofocus required class="form-control" value="<%=(Recordset1.Fields.Item("Peo_password").Value)%>">
                </div>
                <div class="col-lg-10"></div>
                
              <div class="col-lg-10">
                  姓名：  <input name="name" type="text" autofocus required class="form-control" value="<%=(Recordset1.Fields.Item("Peo_name").Value)%>">
                </div>
                <div class="col-lg-10"></div>
              <div class="col-lg-10">
             电话： <input name="phone" type="text" autofocus required class="form-control" value="<%=(Recordset1.Fields.Item("Peo_phone").Value)%>">
                
                <div class="col-lg-10"></div>
                <div class="col-lg-10">
                    <input name="submitt" type="submit" value="修改">
                </div>
                <input type="hidden" name="MM_update" value="form1">
                <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("Peo_no").Value %>">
              </form>
            </div>
        </div>

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
