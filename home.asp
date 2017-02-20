<!DOCTYPE html>
<html>
<head>
    <title>太平人寿风险排查系统</title>
	<meta charset="utf-8" />
	<!-- <link rel="stylesheet" href="/w3.css"> -->
<style>

body {
    font-family: "Lato", sans-serif;
    transition: background-color .5s;
}

.sidenav {
    height: 100%;
    width: 0;
    position: fixed;
    z-index: 1;
    top: 0;
    left: 0;
    background-color: #111;
    overflow-x: hidden;
    transition: 0.5s;
    padding-top: 60px;
}

.sidenav a {
    padding: 8px 8px 8px 32px;
    text-decoration: none;
    font-size: 25px;
    color: #818181;
    display: block;
    transition: 0.3s
}

.sidenav a:hover, .offcanvas a:focus{
    color: #f1f1f1;
}

.sidenav .closebtn {
    position: absolute;
    top: 0;
    right: 25px;
    font-size: 36px;
    margin-left: 50px;
}

#main {	
    transition: margin-left .5s;
    padding: 16px;
}

@media screen and (max-height: 450px) {
  .sidenav {padding-top: 15px;}
  .sidenav a {font-size: 18px;}
}

<!--for search bar   -->
input[type=text] {
    width: 120px;
    box-sizing: border-box;
    border: 2px solid #ccc;
    border-radius: 4px;
    font-size: 16px;
    background-color: white;
    background-image: url('searchicon.png');
    background-position: 10px 10px;
    background-repeat: no-repeat;
    padding: 12px 20px 12px 40px;
    -webkit-transition: width 0.4s ease-in-out;
    transition: width 0.4s ease-in-out;
}

input[type=text]:focus {
    width: 40%;
}
.center {
    margin: auto;
    width: 60%;
    border: 0px solid #73AD21;
    padding: 10px;
}

/* for login small page */

/* Full-width input fields */
input[type=text], input[type=password] {
    width: 30%;
    padding: 1px 14px;
    margin: 8px 0;
    display: inline-block;
    border: 1px solid #ccc;
    box-sizing: border-box;
}

/* Set a style for all buttons */
button {
    background-color: #80bfff;
    color: white;
    padding: 5px 9px;
    margin: 1px 0;
    border: none;
    cursor: pointer;
    width: 30%;
	border-radius: 37px;
    float:right;

  }
}

/* Extra styles for the cancel button */
.cancelbtn {
    width: auto;
    padding: 10px 18px;
    background-color: #f44336;
}

/* Center the image and position the close button */
.imgcontainer {
    text-align: center;
    margin: 24px 0 18px 0;
    position: relative;
}

img.avatar {
    width: 40%;
    border-radius: 60%;
}

.container {
    padding: 40px;
}

span.psw {
    float: right;
    padding-top: 16px;
}

/* The Modal (background) */
.modal {
    display: none; /* Hidden by default */
    position: fixed; /* Stay in place */
    z-index: 1; /* Sit on top */
    left: 0;
    top: 0;
    width: 100%; /* Full width */
    height: 100%; /* Full height */
    overflow: auto; /* Enable scroll if needed */
    background-color: rgb(0,0,0); /* Fallback color */
    background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
    padding-top: 60px;
}

/* Modal Content/Box */
.modal-content {
    background-color: #fefefe;
    margin: 5% auto 15% auto; /* 5% from the top, 15% from the bottom and centered */
    border: 1px solid #888;
    width: 46%; /* Could be more or less, depending on screen size */
}

/* The Close Button (x) */
.close {
    position: absolute;
    right: 25px;
    top: 0;
    color: #000;
    font-size: 35px;
    font-weight: bold;
}

.close:hover,
.close:focus {
    color: red;
    cursor: pointer;
}

/* Add Zoom Animation */
.animate {
    -webkit-animation: animatezoom 0.6s;
    animation: animatezoom 0.3s
}

@-webkit-keyframes animatezoom {
    from {-webkit-transform: scale(0)} 
    to {-webkit-transform: scale(1)}
}
    
@keyframes animatezoom {
    from {transform: scale(0)} 
    to {transform: scale(1)}
}

/* Change styles for span and cancel button on extra small screens */
@media screen and (max-width: 300px) {
    span.psw {
       display: block;
       float: none;
    }
    .cancelbtn {
       width: 100%;
    }	
}

h2 {
  font-size: 45px;
}

h3 {
  font-size: 28px;
}

h2, h3 {
  width: 50%;
  height: 1px;
  margin: 0;
  padding: 0;
  display: inline;
  margin-bottom: 1px !important;
}​

h4{
  width: 50%;
  font-size: 30px;
  height: 0px;
  margin: 0;
  padding: 0;
  display: inline;
  margin-top: -20px !important;
}​

div.small {
    line-height: 2;
}

p.big {
    line-height: 200%;
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

.line {
  fill: none;
  stroke: steelblue;
  stroke-width: 1.5px;
}

</style>


<!--JQuery source used-->
<script src="lib/jquery.min.js"></script>
</head>


<body background="images/homeBackground.jpg" >

<div id="mySidenav" class="sidenav">
  <a href="javascript:void(0)" class="closebtn" onclick="closeNav()">&times;</a>
  <a href="PDFs/fanxiqian.pdf">法律法规</a>
  <a href="PDFs/regulations.pdf">制度规定</a>
  <a href="#" onclick=checkPointClicked("selfEva") >风险点自查</a>
  <a href="#" onclick=checkPointClicked("report") >风险点自查报告</a>
  <a href="#" onclick=checkPointClicked("selfEva") >全面风险排查</a>
  <a href="#" onclick=checkPointClicked("report") >全面风险排查报告</a>
<%
'Get loggin status starts
Dim loggedIn,account,wrongPassWord
loggedIn = "X"
wrongPassWord = "N"
if IsEmpty(Session("LoggedIn")) then
    loggedIn = "N"
else
  if (Session("LoggedIn") = "" ) or ( Session("LoggedIn") = "N" ) then
    loggedIn = "N"
	wrongPassWord = "Y"
  else
	loggedIn=Session("LoggedIn")
	account=Session("setName")
	loggedIn = "Y"
  end if
end If

if loggedIn = "Y" then
	if account = "管理员" then
		response.write 	"    <a href=""admin.asp"">系统管理</a> " & _
						"	</div>							   " & _
						"	<div id=""main"" >				   " & _
						"	<h1 >  </h1>   					   " & _
						"<img src=""images/logo.png"" alt=""Mountain View"" > " & _
						"尊敬的" & account & "用户，您好，"  & "欢迎使用风险排查系统！"  & _
						"<button id=""logOutButton""" & _
						" onclick=""logOut()""" & _  
						" style=""font-size:24px;width:auto;"">退出</button>"
	else 
		response.write 	"									   " & _
						"	</div>							   " & _
						"	<div id=""main"" >				   " & _
						"	<h1 >  </h1>   					   " & _
						"<img src=""images/logo.png"" alt=""Mountain View"" > " & _
						"尊敬的" & account & "用户，您好，"  & "欢迎使用风险排查系统！"  & _
						"<button id=""logOutButton""" & _
						" onclick=""logOut()""" & _  
						" style=""font-size:24px;width:auto;"">退出</button>"
	end if
	'diplay logged in account name
elseif loggedIn = "N" then
	'create logging button 
	response.write "  									   " & _
				   "	</div>							   " & _
				   "	<div id=""main"" >				   " & _
				   "	<h1 >  </h1>   					   " & _
				   "<img src=""images/logo.png"" alt=""Mountain View"" > " & _
				   "<button id=""loginButton""" & _
				   " onclick=""document.getElementById('id01').style.display='block'""" & _  
				   " style=""font-size:24px;width:auto;"">登陆</button>"
else 
	response.write "  									   " & _
				   "	</div>							   " & _
				   "	<div id=""main"" >				   " & _
				   "	<h1 >  </h1>   					   " & _
				   "<img src=""images/logo.png"" alt=""Mountain View"" > " & _
				   "<button id=""loginButton""" & _
				   " onclick=""document.getElementById('id01').style.display='block'""" & _  
				   " style=""font-size:24px;width:auto;"">登陆</button>"
   response.write "登录异常，请联系管理员！"
end if

if wrongPassWord = "Y" then
'response.write "<script> alert(""密码错误，请重新登陆！"")  <script/> "
response.write "<p style=""color:red;"">密码错误，请重新登陆！</p>"
end if


%>
	
<div class="small">
<h6 >  </h6>
<h2>太平人寿保险有限公司</h2>
<h3>甘肃分公司</h3>
<h4>TAIPING LIFE INSURANCE CO.,LTD.GANSU BRANCH</h4>
<hr />​
</div>
<span style="font-size:30px;cursor:pointer" onclick="openNav()">&#9776; 菜单 </span>   
<br>
</div>
</br>
</br>
</br>

<h3>2016年度各部门风险自查违规比例:</h3>
</br>
<svg id="barChartSvg" width="960" height="500"></svg>
</br>
</br>
<h3>公司历年自查情况走势:</h3>
</br>
<svg id="lineChart" width="1300" height="500"></svg>
<script src="lib/d3/d3.js"></script>
<!--bar chart -->
<script>
var svgbarChart = d3.select("#barChartSvg"),
    margin = {top: 20, right: 20, bottom: 30, left: 40},
    width = +svgbarChart.attr("width") - margin.left - margin.right,
    height = +svgbarChart.attr("height") - margin.top - margin.bottom;

var xx = d3.scaleBand().rangeRound([0, width]).padding(0.25),
    y = d3.scaleLinear().rangeRound([height, 0]);

var g = svgbarChart.append("g")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

d3.tsv("violationRatio.tsv", function(d) {
  d.frequency = +d.frequency;
  return d;
}, function(error, data) {
  if (error) throw error;

  xx.domain(data.map(function(d) { return d.letter; }));
  y.domain([0, d3.max(data, function(d) { return d.frequency; })]);

  g.append("g")
      .attr("class", "axis axis--x")
      .attr("transform", "translate(0," + height + ")")
      .call(d3.axisBottom(xx));

  g.append("g")
      .attr("class", "axis axis--y")
      .call(d3.axisLeft(y).ticks(5, "%"))
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
      .attr("x", function(d) { return xx(d.letter); })
      .attr("y", function(d) { return y(d.frequency); })
      .attr("width", xx.bandwidth())
      .attr("height", function(d) { return height - y(d.frequency); });
});

<!--line chart -->
var svgLineChart = d3.select("#lineChart"),
    margin = {top: 20, right: 20, bottom: 30, left: 50},
    width = +svgLineChart.attr("width") - margin.left - margin.right,
    height = +svgLineChart.attr("height") - margin.top - margin.bottom,
    gg = svgLineChart.append("g").attr("transform", "translate(" + margin.left + "," + margin.top + ")");

var parseTime = d3.timeParse("%d-%b-%y");

var x = d3.scaleTime()
    .rangeRound([0, width]);

var y = d3.scaleLinear()
    .rangeRound([height, 0]);

var line = d3.line()
    .x(function(d) { return x(d.date); })
    .y(function(d) { return y(d.close); });

d3.tsv("trendLineBar.tsv", function(d) {
  d.date = parseTime(d.date);
  d.close = +d.close;
  return d;
}, function(error, data) {
  if (error) throw error;

  x.domain(d3.extent(data, function(d) { return d.date; }));
  y.domain(d3.extent(data, function(d) { return d.close; }));

  gg.append("g")
      .attr("class", "axis axis--x")
      .attr("transform", "translate(0," + height + ")")
      .call(d3.axisBottom(x));

  gg.append("g")
      .attr("class", "axis axis--y")
      .call(d3.axisLeft(y))
    .append("text")
      .attr("fill", "#000")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", "0.71em")
      .style("text-anchor", "end")
      .text("违规数");

  gg.append("path")
      .datum(data)
      .attr("class", "line")
      .attr("d", line);
});

</script>

<div id="id01" class="modal">
  <form class="modal-content animate" action="login.asp">
    <div class="imgcontainer">
      <span onclick="document.getElementById('id01').style.display='none'" class="close" title="Close Modal">&times;</span>
      <img src="images/avatar.png" alt="Avatar" class="avatar">
    </div>

    <div class="container">
      <label><b>账号</b></label>
      <input type="text" placeholder="输入账号" name="userName" required>
      <label><b>密码</b></label>
      <input type="password" placeholder="输入密码" name="passWord" required>
      <button type="submit" style="font-size:20px;">登陆</button>  
    </div>
  </form>
</div>
<script>
function logOut(){
	//navigation("logOut.asp")
	var xmlhttp = new XMLHttpRequest();
	xmlhttp.onreadystatechange = function() {
		if (this.readyState == 4 && this.status == 200) {
			//loggedIn = this.responseText;
			window.location.reload()
		}
	}
	xmlhttp.open("GET", "logOut.asp", true); 
	xmlhttp.send();
}
function checkPointClicked(navType){
	var navigateTo
	//starts, get logged In status 
	var loggedIn = "N";
	var xmlhttp = new XMLHttpRequest();
	xmlhttp.onreadystatechange = function() {
		if (this.readyState == 4 && this.status == 200) {
			loggedIn = this.responseText;
				if (loggedIn == 'Y'){
					navigateTo="/selfEvaNavigation.asp?navType=" + navType;
					navigation(navigateTo);
				} else {
					document.getElementById('id01').style.display='block'
				}
		}
	}
	xmlhttp.open("GET", "loginChecker.asp", true); 
	xmlhttp.send();
}

function navigation(navigateTo){
 window.location.href = navigateTo
}; 

function turnOfff(id) {
	document.write("id is :"+ id);
    document.getElementById("mySidenav").style.width = "250px";
    document.getElementById("main").style.marginLeft = "250px";
    document.body.style.backgroundColor = "rgba(0,0,0,0.4)";
}

function turnOff() {
    document.getElementById("mySidenav").style.width = "250px";
    document.getElementById("main").style.marginLeft = "250px";
    document.body.style.backgroundColor = "rgba(0,0,0,0.4)";
}

function openNav() {
    document.getElementById("mySidenav").style.width = "250px";
    document.getElementById("main").style.marginLeft = "250px";
    document.body.style.backgroundColor = "rgba(0,0,0,0.4)";
}

function closeNav() {
    //document.getElementById("mySidenav").hide();
    document.getElementById("mySidenav").style.width = "0";
    document.getElementById("main").style.marginLeft= "0";
    document.body.style.backgroundColor = "white";
}
</script>
<!-- <footer>太平人寿保险有限公司</footer> -->
</body>
</html>
