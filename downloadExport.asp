<% 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=自查报告.xls"
Dim sConnection, objConn , objRS ,headerRow, queryStr ,totalCheckPoints,subCat
'subCat = Request.querystring("subCat")
subCat=1
sqlGetCount="select count(1) totalRecord from webone.reporting where sub_cat_id  = " & "'" & subCat & "' ;"
sqlGetCount="select count(1) totalRecord from webone.reporting where 1=1;"
queryStr="SELECT set_name , sub_cat_name , checkpoint , issue,corrections, result, audit_id,sub_cat_id,checkpoint_id FROM webone.reporting where 1= 1  "
queryStr=queryStr & " AND sub_cat_id  = " & "'" & subCat & "'"  & " order by  audit_id ASC ;"
queryStr="SELECT set_name , sub_cat_name , checkpoint , issue,corrections, result, audit_id,sub_cat_id,checkpoint_id FROM webone.reporting where 1= 1  ;"
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(sqlGetCount)



While Not objRS.EOF
totalCheckPoints=objRS.Fields("totalRecord")
objRS.MoveNext
Wend
Set objRS = objConn.Execute(queryStr)
'preparing search result buttons
%>
<table  BORDER=1>
<tr> <td colspan="7"> 自查报告  </td> </tr>
<tr> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>排查发现问题</th>  <th>处置整改措施</th>  <th>整改完成情况</th>  <th>自查时间</th> </tr>

<%
Dim checkPointIndex	
checkPointIndex=1
While Not objRS.EOF
IF checkPointIndex=1 Then
%>
<tr> <td rowspan="<% = totalCheckPoints %>">  <% = objRS.Fields("set_name") %> </td>
<td rowspan="<% = totalCheckPoints %>"> <% = objRS.Fields("sub_cat_name") %>  </td> 
<td >  <% =  objRS.Fields("checkpoint") %> </td> 
<td > <% = objRS.Fields("issue") %> </td>
<td > <% = objRS.Fields("corrections") %> </td>
<td > <% = objRS.Fields("result")  %> </td>
<td >'<% =  objRS.Fields("audit_id") %></td> </tr>
<% End If 

IF checkPointIndex>1 Then
%>
<tr > <td>  <% =  objRS.Fields("checkpoint") %> </td>
<td>  <% =  objRS.Fields("issue") %> </td>
<td>  <% = objRS.Fields("corrections") %> </td>
<td>  <% = objRS.Fields("result") %> </td>
<td>'<% =  objRS.Fields("audit_id") %></td> </tr>

<%
End If
checkPointIndex=checkPointIndex+1
objRS.MoveNext
Wend
%>
</table>
<%
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>