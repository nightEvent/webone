<%
Response.CharSet = "utf-8"
response.expires=-1
Session.Timeout=1440

Dim sConnection, objConn , objRS , queryStr ,checkPointsList
checkPointsList = Request.Form("checkPointIdList")

queryStr="DELETE FROM webone.checkpoints WHERE checkpoint_id in (" & checkPointsList & ");"

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
Response.Write "选择的项目已删除！"
%>
