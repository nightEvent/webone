<%
response.expires=-1
Dim sConnection, objConn , objRS ,headerRow, queryStr
headerRow="结果如下..."
queryStr="SELECT CONVERT(sub_cat_id USING utf8) sub_cat_id_str, set_name, category_name, sub_cat_name FROM subCat where 1= 1 AND ( "
queryStr=queryStr & " set_name LIKE " & "'" & "%" & request.querystring("q") & "%" & "'" 
'queryStr=queryStr & " OR category_name LIKE " & "'" & "%" & request.querystring("q") & "%" & "'" 
'queryStr=queryStr & " contect LIKE " & "'" & "%" & request.querystring("q") & "%" & "'" 
'queryStr=queryStr & " rectification LIKE " & "'" & "%" & request.querystring("q") & "%" & "'" 
queryStr=queryStr & " );"


sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)


'preparing search result buttons
Response.Write headerRow &  "<br>" 
While Not objRS.EOF
'Response.Write "sub_cat_id_str is A :" & objRS.Fields("sub_cat_id_str") &  "sub_cat_id_str is B :"  & objRS.Fields("sub_cat_id_str") & "sub_cat_id_str is C :"  & objRS.Fields("sub_cat_id_str") &  "<br>"

'Response.Write  objRS.Fields("set_name") &  "   "  & objRS.Fields("category_name")  &  "   " & objRS.Fields("sub_cat_name")  &  "<br>"
'Response.Write  "<button id=subCat" & objRS.Fields("sub_cat_id_str") & " onclick=" & """lightOn("   & objRS.Fields("sub_cat_id_str")  &")""" & "> " & objRS.Fields("set_name") &  "   "  & objRS.Fields("category_name")  &  "   " & objRS.Fields("sub_cat_name")  & "</button>" &  "<br>"
Response.Write  "<button id=subCat" & objRS.Fields("sub_cat_id_str") & " onclick=" & """lightOn(" & ")""" & "> " & objRS.Fields("set_name") &  "   "  & objRS.Fields("category_name")  &  "   " & objRS.Fields("sub_cat_name")  & "</button>" &  "<br>"
objRS.MoveNext
Wend

objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>