<%
Response.CharSet = "utf-8"
Session.Timeout=1440
response.expires=-1
Dim sConnection, objConn ,objRS, queryStr, userId,Newpassword
userId=Request.Form("userId")
Newpassword=Request.Form("newPassword")
'Response.Write "userId:" & userId & " Newpassword:" & Newpassword
queryStr="UPDATE webone.user set passwordHash =  UNHEX(SHA1('" & Newpassword & "')) " & " where user_id =  " &  userId & " ;"
'queryStr="SELECT 1 FROM DUAL;"
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection)
Set objRS = objConn.Execute(queryStr)
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
Response.Write "Y"
%>