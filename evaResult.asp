<!DOCTYPE html>
<html>
<head>
    <title>Hi There</title>
	<meta charset="utf-8" />
	<!-- <link rel="stylesheet" href="/w3.css"> -->
<style>

</style>

<!--JQuery source used-->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>


</head>


<body>



<!DOCTYPE html>
<html>
<head>
	<style>
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

	</style>
</head>
<body>

<p>.</p>

<!--<table class="editabletable"> -->
<table >
  <tr>
  <td colspan="7">
     甘肃分公司内控风险自查表
  </td>
  </tr>
  <tr>
    <th>系列</th>
    <td>分类</td>
    <td>详细风险点</td>
    <td>合规要求</td>
    <td>排查方法</td>
    <td>排查经过</td>
    <td>是否发现问题</td>
  </tr>

  <tr>
	<td rowspan="6" > 个险 </td>
	<td rowspan="6" > 入、离司资料不合规 </td>
  </tr>
  
  <tr>  
       <td >详细风险点一</td>
       <td >合规要求一</td>
	   <td rowspan="5" >按照相关制度要求逐一排查业务人员入司档案，要求上报资料是否齐全，签字是否齐全，是否核实三无人员的具体情况，核查系统上岗人员是否持有执业证，执业证申请时间是否在入司之前。</td>
	   <!-- <td contenteditable></td>   跳过排查方法--> 
	   <td  contenteditable>排查经过一</td> 
	   <td contenteditable>是否发现问题一</td> 
  </tr>
  
  <tr>  
     <td contenteditable>详细风险点二</td>  
	 <td contenteditable>合规要求二</td>
     <!-- <td contenteditable></td>   跳过排查方法--> 
	 <td  contenteditable>排查经过二</td> 
	 <td contenteditable>是否发现问题二</td> 
  </tr>
  <tr> 
     <td contenteditable>详细风险点三</td>  
	 <td contenteditable>合规要求三</td>
     <!-- <td contenteditable></td>   跳过排查方法--> 
	 <td  contenteditable>排查经过三</td> 
	 <td contenteditable>是否发现问题三</td> 
  </tr>
  <tr> 
     <td contenteditable>详细风险点四</td>  
	 <td contenteditable>合规要求四</td>
     <!-- <td contenteditable></td>   跳过排查方法--> 
	 <td  contenteditable>排查经过四</td> 
	 <td contenteditable>是否发现问题四</td> 
  </tr>
</table>



<div id="checkPointArea">
<button type="button" onclick="loadDoc()">Change Content</button>
</div>
<textarea id=textarea  rows="4" cols="50">
At w3schools.com you will learn how to make a website. We offer free tutorials in all web development technologies.
</textarea>



</body>


</html>
