<%
response.expires=-1
Dim sConnection, objConn , objRS ,headerRow, queryStr 
subCat = Request.querystring("subCat")
'subCat=(len(subCat) - len(replace(subCat, ",", "")) + 1)


'queryStr="SELECT CONVERT(sub_cat_id USING utf8) sub_cat_id, set_name , category_name , checkpoint , fulfill_standard, audit_rule, sub_cat_name,checkpoint_id FROM webone.selfEva where 1= 1  "
'queryStr=queryStr & " AND sub_cat_id  = " & "'" & request.querystring("q") & "'"  & " order by  checkpoint_id ASC ;"
queryStr="SELECT CONVERT(sub_cat_id USING utf8) sub_cat_id, set_name , category_name , checkpoint , fulfill_standard, audit_rule, sub_cat_name,checkpoint_id FROM webone.selfEva where 1= 1  "
queryStr=queryStr & " AND sub_cat_id  = " & "'" & subCat & "'"  & " order by  checkpoint_id ASC ;"

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)


'preparing search result buttons
Dim topHead, secondH, count,startCheck, buttonSubmit
topHeader=" <tr> <td colspan=""7""> 甘肃分公司内控风险自查表  </td> </tr> "
secondH=" <tr> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>合规要求</th>  <th>排查方法</th>  <th>排查经过</th> <th>发现问题</th>  </tr> "
count="5"
startCheck="<button onclick=""startChecking()""> 排查开始 </button>"
buttonSubmit="<button onclick=""startChecking()""> 填写完成进入下一页 </button>"
Response.Write "<table id=""selfEvaSheet"" >"
Response.Write topHeader
Response.Write secondH
Dim checkPointsCount
checkPointsCount=1
While Not objRS.EOF
IF checkPointsCount=1 Then
Response.Write "<tr> <td rowspan=""" & count & """ contenteditable >"  & objRS.Fields("set_name")      &  "</td> "
Response.Write "<td rowspan="""      & count & """ contenteditable >"  & objRS.Fields("category_name") &  "</td>  "

Response.Write " <td  contenteditable >" & objRS.Fields("checkpoint") &  "</td>  <td  contenteditable > " & objRS.Fields("fulfill_standard") & "</td> "

Response.Write "<td rowspan=""" & count & """ contenteditable>"  & objRS.Fields("audit_rule") & "</td>"

Response.Write "<td contenteditable >排查经过</td> " 'when row number is 2, read from cell NUMBER 5 through 8, when cell number is 6 , it's a checkbox

Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" checked>  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"

End If

IF checkPointsCount>1 Then 'starting from number 3, read from cell number 2 through 5 and when cell number is 3 it's a checkbox
Response.Write "<tr> <td  contenteditable >" & objRS.Fields("checkpoint") &  "</td>  <td  contenteditable > " & objRS.Fields("fulfill_standard") & "</td> "
Response.Write "<td  contenteditable>排查经过..</td> " 
Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" checked>  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
End If
checkPointsCount=checkPointsCount+1
objRS.MoveNext
Wend
Response.Write "</table>"
Response.Write "<br><br><br>"
'Response.Write startCheck 
Response.Write buttonSubmit
Response.Write "<input type=""hidden"" name=""orderNumber"" id=""checkPointsCount"" value=""" & ( checkPointsCount - 1 ) & """ />"
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>