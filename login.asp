<%
Session.Timeout=1440
response.expires=-1
Dim account, passWord, loggedIn,previousPage
Session("loggedIn")="N"
account = Request.querystring("account")
passWord = Request.querystring("passWord")
passWord = Trim(passWord)
Dim sConnection, objConn , objRS ,headerRow, queryStr, hostname
queryStr=" SELECT set_name,account FROM webone.user where account='" & account  & "' and passwordHash=UNHEX(SHA1('" & passWord & "')) ;" 
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
While Not objRS.EOF
Session("LoggedIn") = "Y"
Session("setName") = objRS.Fields("set_name")
Session("account") = objRS.Fields("account")
objRS.MoveNext
Wend
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
previousPage = Request.ServerVariables("HTTP_REFERER")
Response.Redirect previousPage
%>