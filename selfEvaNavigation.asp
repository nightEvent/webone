<!DOCTYPE html>
<html>
<header>
    <title>太平人寿风险排查系统</title>
	<meta charset="utf-8" />
<style>
a, u {
			text-decoration: none;
		}
details summary p.small {
    line-height: 10%;
	
}
details {
    border-radius: 3px;
    background: #EEE;
}
details summary {
    font-size: 27px;
    vertical-align: top;
    background: #333;
    color: #FFF;
    border-radius: 3px;
    padding: 5px 10px;
    outline: none;
}
button {
    background-color: #1e90ff;
    color: white;
    padding: 1px 1px;
    margin: 1px 0;
    border: none;
    cursor: pointer;
    width: 3%;
	border-radius: 80px;
    float:left;
  }


a.hoverEff:hover,a.hoverEff.active { 
    //text-decoration: underline;
	font-size: 150%;
	//color: #b3ffb3;
 }
 
 
a.homeLink:hover,a.homeLink.active{
  text-decoration: underline;
}
 
.bar {
  fill: steelblue;
}

.bar:hover {
  fill: brown;
}

.axis--x path {
  display: none;
}
</style>
</header>

<body  background="images/singleCatSheet.jpg">
<a class="homeLink" href="http://localhost:88/home.asp">Home</a>
<br>
<h2>系列</h2>
<%
Session.Timeout=60
response.expires=-1
Response.CharSet = "utf-8"
Dim sConnection, objConn , objRS ,headerRow, queryStr, hostname,setName,reqType,chkType
hostname="localhost:88"
setName=Session("setName")
reqType=Request.querystring("reqType")
chkType=Request.querystring("chkType")
if setName = "管理员" then
queryStr="SELECT distinct sub_cat_id, set_name,sub_cat_name FROM webone.subCatNav where 1= 1 " & _
         " AND  chk_Type = '" &  chkType & "' ;"
else
'queryStr="SELECT distinct sub_cat_id, set_name,sub_cat_name FROM webone.subCatNav where 1= 1  and set_name = """  & setName & """ ;"
queryStr="SELECT distinct sub_cat_id, set_name,sub_cat_name FROM webone.subCatNav where 1= 1  " & _ 
         " AND set_name = """  & setName & """"  & _
         " AND chk_Type ='" & chkType & "' ;"
end if

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
Dim previous_set, current_set,index
previous_set="enjoyLife"
index=1

While Not objRS.EOF
current_set=objRS.Fields("set_name")
if current_set  <> previous_set then
 if index <> 1 then
  Response.Write " </details>"
 end if
  Response.Write " <details>   <summary> " & current_set & "</summary>"
end if
if reqType = "selfEva" then
Response.Write " <a href=""" & "singleCatSheet.asp?subCatId=" & objRS.Fields("sub_cat_id") & """>" & objRS.Fields("sub_cat_name")  & "</a> <br>"  
else
Response.Write " <a href=""" & "issueReport.asp?subCatId=" & objRS.Fields("sub_cat_id") & """>" & objRS.Fields("sub_cat_name")  & "</a> <br>"  
end if
previous_set = current_set
index=index+1
objRS.MoveNext
Wend
Response.Write " </details>"
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
Response.Write "</br>"
if reqType <> "selfEva" and setName <> "" then  'setName <> "" is to avoid session expire
Response.Write "<a class=""hoverEff""  href=""downloadExport.asp"">生成表格并下载</a>"
Response.Write "</br>"
Response.Write "<a class=""hoverEff"" href=""docReportBuilder.asp"">生成报告并下载</a>"
Response.Write "</br>"
Response.Write "<a class=""hoverEff""  href=""#"" onclick=""buildBarChart()"">2016违规比例</a>"
Response.Write "<a class=""hoverEff""  href=""createMultipleSheetsTest.asp"" >多表Excel下载</a>"
end if
%>

<h4 id=title></h4>
<svg width="600" height="300"></svg>
<script src="d3/d3.js"></script>
<script>
function buildBarChart(){

document.getElementById("title").innerHTML = "<h3>2016年度各部门风险自查违规比例:</h3>"


var svg = d3.select("svg"),
    margin = {top: 20, right: 20, bottom: 30, left: 40},
    width = +svg.attr("width") - margin.left - margin.right,
    height = +svg.attr("height") - margin.top - margin.bottom;

var x = d3.scaleBand().rangeRound([0, width]).padding(0.1),
    y = d3.scaleLinear().rangeRound([height, 0]);

var g = svg.append("g")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

d3.tsv("violationRatio.tsv", function(d) {
  d.frequency = +d.frequency;
  return d;
}, function(error, data) {
  if (error) throw error;

  x.domain(data.map(function(d) { return d.letter; }));
  y.domain([0, d3.max(data, function(d) { return d.frequency; })]);

  g.append("g")
      .attr("class", "axis axis--x")
      .attr("transform", "translate(0," + height + ")")
      .call(d3.axisBottom(x));

  g.append("g")
      .attr("class", "axis axis--y")
      .call(d3.axisLeft(y).ticks(10, "%"))
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", "0.71em")
      .attr("text-anchor", "end")
      .text("Frequency");

  g.selectAll(".bar")
    .data(data)
    .enter().append("rect")
      .attr("class", "bar")
      .attr("x", function(d) { return x(d.letter); })
      .attr("y", function(d) { return y(d.frequency); })
      .attr("width", x.bandwidth())
      .attr("height", function(d) { return height - y(d.frequency); });
});
}
</script>


<br> <br>


<p><b></b></p>

</body>
</html>