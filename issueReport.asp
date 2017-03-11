<!DOCTYPE html>
<html>
<%@ language="VBScript" codepage="65001" %>
<%
Response.CharSet = "utf-8"
session.codepage=65001
%>

<head>
<title>问题报告</title>
<link rel="icon" href="images/favicon.ico" type="image/x-icon">

	<style>
		a, u {
			text-decoration: none;
		}
		table, td, th {
			border: 1px solid black;
		}

		table {
			border-collapse: collapse;
			width: 100%;

		}

		th {
			height: 50px; vertical-align:center;text-align:left;
		}
		td {vertical-align:center;text-align:center;}
div.fixed {
    position: fixed;
 
    right: 1;
    width: 300px;
   
       }

	   
button {
    background-color: #A9A9A9;
    color: white;
    padding: 1px 1px;
    margin: 1px 0;
    border: none;
    cursor: pointer;
    width: 3%;
	border-radius: 80px;
    float:left;
  }

</style>

</head>

<body background="images/singleCatSheet.jpg">

<script>

</script>

<a href="http://localhost:88/home.asp">Home</a> <br> <br>
<%
response.expires=-1
Dim sConnection, objConn , objRS ,headerRow, queryStr ,totalCheckPoints
subCat = Request.querystring("subCat")

sqlGetCount="select count(1) totalRecord from webone.reporting where sub_cat_id  = " & "'" & subCat & "' ;"
queryStr="SELECT set_name , sub_cat_name , checkpoint , issue,corrections, result, audit_id,sub_cat_id,checkpoint_id FROM webone.reporting where 1= 1  "
queryStr=queryStr & " AND sub_cat_id  = " & "'" & subCat & "'"  & " order by  audit_id ASC ;"

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(sqlGetCount)
While Not objRS.EOF
totalCheckPoints=objRS.Fields("totalRecord")
response.write "自查发现问题数目 ： " & totalCheckPoints
objRS.MoveNext
Wend
Set objRS = objConn.Execute(queryStr)
'preparing search result buttons
Dim topHead, secondH, buttonSubmit
topHeader=" <tr> <td colspan=""7""> 自查报告  </td> </tr> "
secondH=" <tr> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>排查发现问题</th>  <th>处置整改措施</th>  <th>整改完成情况</th>  <th>自查时间</th> </tr> "


Response.Write "<table id=""selfEvaSheet"" >"
Response.Write topHeader
Response.Write secondH
Dim checkPointIndex	
checkPointIndex=1
While Not objRS.EOF
IF checkPointIndex=1 Then
Response.Write "<tr> <td rowspan=""" & totalCheckPoints & """  >"  & objRS.Fields("set_name")      &  "</td> "
Response.Write "<td rowspan="""      & totalCheckPoints & """ >"  & objRS.Fields("sub_cat_name") &  "</td>  "
Response.Write "  <td   >" &  " <a href=""#"" onclick=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")"" >"  &  objRS.Fields("checkpoint") &  "</a> </td>" 
Response.Write "<td >" & objRS.Fields("issue") &  "</td> "
Response.Write "<td >" & objRS.Fields("corrections") &  "</td> "
Response.Write "<td >" & objRS.Fields("result") &  "</td> "
Response.Write "<td >" & objRS.Fields("audit_id") &  "</td> </tr> "
End If

IF checkPointIndex>1 Then
Response.Write "<tr > <td> <a href=""#"" onclick=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")"" >" & objRS.Fields("checkpoint") &  "</a> </td> "
Response.Write "<td >" & objRS.Fields("issue") & "</td> "
Response.Write "<td >" & objRS.Fields("corrections") &  "</td> "
Response.Write "<td >" & objRS.Fields("result") &  "</td>  "
Response.Write "<td >" & objRS.Fields("audit_id") &  "</td>  </tr>"
End If
checkPointIndex=checkPointIndex+1
objRS.MoveNext
Wend
Response.Write "</table>"
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
<br>
&nbsp <button onclick="buttonBack()"  style="font-size:17px;" > 返回 </button>
<p id="linkedRules"></p>

<script>
function checkPointClicked(checkPointId){
   document.getElementById("linkedRules").style.border = "thin solid black"
    if (checkPointId.length == 0) { 
        document.getElementById("linkedRules").innerHTML = "  ";
        return;
    } else {
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (this.readyState == 4 && this.status == 200) {
                document.getElementById("linkedRules").innerHTML = this.responseText;
            }
        };
        xmlhttp.open("GET", "linkedRules.asp?q="+checkPointId, true); 
        xmlhttp.send();
    }
}

function navigates(navigateTo){
 window.location.href = navigateTo
}; 

function buttonBack(){
navigates(document.referrer)
}
</script>
</body>
</html>