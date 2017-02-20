<%
Session.Timeout=1440
response.expires=-1
Dim userName, passWord, loggedIn,previousPage
Session("loggedIn")="N"
userName = Request.querystring("userName")
passWord = Request.querystring("passWord")
'previousPage = Request.querystring("previousPage")
'previousPage
'Response.Write "previousPage is " & previousPage

Dim sConnection, objConn , objRS ,headerRow, queryStr, hostname
queryStr=" SELECT set_name FROM webone.user where account='" & userName  & "' and password='" & passWord & "' ;"
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
While Not objRS.EOF
Session("LoggedIn") = "Y"
Session("setName") = objRS.Fields("set_name")
'response.write= "alert(""" & objRS.Fields("set_name") & """)"
objRS.MoveNext
Wend
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
previousPage = Request.ServerVariables("HTTP_REFERER")
Response.Redirect previousPage
%>