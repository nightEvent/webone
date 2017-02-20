<%
response.expires=-1
Dim sConnection, objConn , objRS ,headerRow, queryStr 
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
checkPointId = Request.querystring("q")
'checkPointId=(len(checkPointId) - len(replace(checkPointId, ",", "")) + 1)

queryStr="SELECT checkpoint_id, regulation ,law FROM webone.linkedLawsNdReg where 1= 1  "
queryStr=queryStr & " AND checkpoint_id  = "  & checkPointId  & " order by  checkpoint_id ASC ;"


Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
Dim ind
ind=1
While Not objRS.EOF
response.write  "对应的制度及条文 " & ind & "</br>"  & "制度 ： </br> " &  objRS.Fields("regulation") & "</br>"   & "条文 ：</br>" & objRS.Fields("law")
ind= ind + 1
objRS.MoveNext
Wend

objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>