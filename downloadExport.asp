<% 
Response.CharSet = "utf-8"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=历史排查数据.xls"
Dim sConnection, objConn , objRS ,headerRow, queryStr ,totalCheckPoints,pastNdays,timeLowest
pastNdays = Request.querystring("pastNdays")
pastNdays= -1 * pastNdays
'sqlGetCount="select count(1) totalRecord from webone.reporting where sub_cat_id  = " & "'" & subCat & "' ;"
'sqlGetCount="select count(1) totalRecord from webone.reporting where 1=1;"
queryStr="SELECT set_name , sub_cat_name ,chk_Type, checkpoint , issue,corrections, result, creation_time,procedures,sub_cat_id,checkpoint_id FROM webone.reporting where 1= 1  "
queryStr=queryStr & " AND creation_time  >  TIMESTAMPADD(DAY, " & pastNdays  &  ",  Now() ) " 
queryStr=queryStr & " ORDER BY chk_Type, creation_time ;"
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
'preparing search result buttons
%>
<table  BORDER=1>
<tr> <td colspan="7"> 自查报告  </td> </tr>
<tr> <th>排查类型</th> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>排查经过</th>  <th>排查发现问题</th>  <th>处置整改措施</th>  <th>整改完成情况</th>  <th>自查时间</th>  </tr>

<%
Dim checkPointIndex,auditTime
While Not objRS.EOF
%>
<tr> <td >  <% = objRS.Fields("chk_Type") %> </td>
<td >  <% = objRS.Fields("set_name") %> </td>
<td > <% = objRS.Fields("sub_cat_name") %>  </td> 
<td >  <% =  objRS.Fields("checkpoint") %> </td> 
<td >  <% =  objRS.Fields("procedures") %> </td> 
<td > <% = objRS.Fields("issue") %> </td>
<td > <% = objRS.Fields("corrections") %> </td>
<td > <% = objRS.Fields("result")  %> </td>
<td >'<% =  objRS.Fields("creation_time") %></td> </tr>

<%
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