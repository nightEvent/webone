<%
response.expires=-1
'sqlStr="SELECT content, rectification FROM checkpoints WHERE sub_category_id =  "
sqlStr="SELECT content, rectification FROM webone.checkpoints WHERE sub_category_id = "   
sqlStr=sqlStr & "'" & request.querystring("q") & "'"

Dim mysqlConnection, objConn , RS 

mysqlConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(mysqlConnection) 
Set RS = objConn.Execute(sqlStr)

response.write "请点下一条开始自查.."
do until RS.EOF
  for each x in RS.Fields
  'response.write RS.Fields("content")  & RS.Fields("rectification")
  
  next
  RS.MoveNext
loop
%>