<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/chris.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_chris_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Bow3 WHERE Bow3_peo in(SELECT * FROM dbo.login)" 
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
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
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

<body background="background/BingWallpaper-2017-06-02.jpg">
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
        <li><a href="About.html">关于</a></li> 
        <li><a href="TheFirstPage1.asp">借用</a></li>
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
<form action="TheFirstPage.asp" method="get"></form>
<table class="table table-striped" style="margin-left:50px;margin-right:150px; background-color:#F0EBEB; width:1260px">
	<caption style="font-size:24px; background-color:#F0EBEB; "><strong>借用记录:</strong></caption>
        <thead>
            <tr>
                <th>日期</th>  
                <th>时间（单位：节）</th> 
                <th>教室</th> 
                <th>事由</th> 
                <th>借用人</th> 
          </tr>  
        </thead>  
		<tbody>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <th><%=(Recordset1.Fields.Item("Bow3_data").Value)%> </th>
    <th><%=(Recordset1.Fields.Item("Bow3_time").Value)%></th>
    <th><%=(Recordset1.Fields.Item("Bow3_classno").Value)%> </th>
    <th><%=(Recordset1.Fields.Item("Bow3_reason").Value)%></th>
    <th><%=(Recordset1.Fields.Item("Bow3_peo").Value)%></th>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
        </tbody>  
</table>
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
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
