<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/chris.asp" -->
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
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>教室借用系统</title>

<!-- Bootstrap -->
<link href="CSS/bootstrap.css" rel="stylesheet">


</head>

<body background="background/BingWallpaper-2017-05-23.jpg">
<!--引入js文件-->
<script src="js/jquery.min.js"></script>
<script src="js/dropdown.js"></script>
<script src="js/bootstrap.js"></script>
<script src="js/table.js"></script>
<!--导航栏-->
<nav class="navbar navbar-default" >
<div class="container-fluid"> 
  <div class="collapse navbar-collapse" id="defaultNavbar1">
      <ul class="nav navbar-nav">
        <li><a href="Login.asp">注销</a></li>
        <li><a href="manager.asp">审核</a></li>
        
      </ul>
      <div class="container-fluid">
  		<div class="row">
    		<div class="col-md-offset-0 col-md-7">
      			<h3 class="text-center">教室借用系统</h3>
    		</div>
			</div>
		</div>
 	</div>
  	</div>
</nav>
 <br>
<form name="form1">

  <table class="table table-striped" style="margin-left:50px;margin-right:150px; background-color:#F0EBEB; width:1260px">
    <caption style="font-size:24px; background-color:#F0EBEB; text-align:center "><strong>用户信息一览表</strong></caption>
    <thead>
      <tr>
        <th>用户账号</th>  
        <th>用户姓名</th> 
        <th>联系方式</th> 
       
      </tr>  
      </thead>  
    <tbody>
      <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
        <tr>
          <th><%=(Recordset1.Fields.Item("Peo_no").Value)%></th>
          <th><%=(Recordset1.Fields.Item("Peo_name").Value)%></th>
          <th><%=(Recordset1.Fields.Item("Peo_phone").Value)%> </th>
         
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
      </tbody>  
</table></form>
<nav class="navbar navbar-default navbar-fixed-bottom">
    <div class="navbar-inner navbar-content-center">
    <p class="text-center" style="padding: 5px;font-size:20px">Author@1122</p>
    </div>
   </nav>
   </body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
