<!DOCTYPE html>
<html>
<head>
<link rel="shortcut icon" href="favicon.ico" >
    <title>太平人寿风险排查系统</title>
	<meta charset="utf-8" />
	<!-- <link rel="stylesheet" href="/w3.css"> -->
<style>

a.homeLink:hover,a.homeLink.active{
  text-decoration: underline;
}
</style>
<!--JQuery source used-->
<script src="lib/jquery.min.js"></script>
</head>
<body background="images/homeBackground.jpg" >
<a class="homeLink" href="http://localhost:88/home.asp">主页</a>  &nbsp  &nbsp <a class="homeLink" href=" admin.asp">返回系统管理页面</a> <br> <br>
<label>用户</label> <br>
<select id="userIdSelected" name="用户" autofocus>
<option value="XX" style="display:none;"></option>
	<%
		Dim sConnection, objConn , objRS ,headerRow, queryStr, hostname
		queryStr=" select user_id, account from webone.user ;"
		sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
		Set objConn = Server.CreateObject("ADODB.Connection") 
		objConn.Open(sConnection) 
		Set objRS = objConn.Execute(queryStr)
		While Not objRS.EOF
	%>
		<option value="<%=objRS.Fields("user_id")%>"><%=objRS.Fields("account")%></option>
	<%
		objRS.MoveNext
		Wend
		objRS.Close
		Set objRS = Nothing
		objConn.Close
		Set objConn = Nothing
	%>
</select>
<br><br>
<label>输入新密码</label> <br> 
<input id="newPassword" type="text"></input> 
<button onclick="changePassWord()" type="input">提交</button>
<script>
function changePassWord(){
	var userIdSelected 	= document.getElementById("userIdSelected");
	var newPassword 	= document.getElementById("newPassword").value;
	var userId          = userIdSelected.options[ userIdSelected.selectedIndex ].value;
	var params			= "userId=" + userId + "&newPassword=" + newPassword; 
	var http = new XMLHttpRequest();
	http.open("POST", "passwordResetServer.asp", true);
	//Send the proper header information along with the request
	http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	http.onreadystatechange = function(){//Call a function when the state changes.
		if(http.readyState == 4 && http.status == 200){
		  var responseText = http.responseText;
		  if (responseText=="Y"){
		      alert("密码修改成功！");
		    }
		  location.reload(true);
		 }
	   }
	http.send(params)
}
</script>
</body>
</html>
