<!DOCTYPE html>
<html>
<head>
    <title>添加风险点</title>
	<link rel="icon" href="images/favicon.ico" type="image/x-icon">
	<meta charset="utf-8" />
	<!-- <link rel="stylesheet" href="/w3.css"> -->
<style>
a.homeLink:hover,a.homeLink.active{
  text-decoration: underline;
}

.hiding {
	visibility: hidden;
}

.displaying {
	visibility: visible;
}
</style>
<!--JQuery source used-->
<script src="lib/jquery.min.js"></script>
</head>


<body >
<a class="homeLink" href="home.asp">Home</a> <br> <br>
<!-- <body background="images/homeBackground.jpg" > -->

<label>选择添加类型:</label>
<br>
<input id="chckTypeZ" type="radio" onclick="chckTypeChanged('Z')" name="chckType" value="Z" checked> 自查
<input id="chckTypeL" type="radio" onclick="chckTypeChanged('L')" name="chckType" value="L"> 临时性检查
<input id="chckTypeQ" type="radio" onclick="chckTypeChanged('Q')" name="chckType" value="Q"> 全面排查
<br>

<div id="adminPageData" data_checkpoint_type="Z" ></div>

<br>
<div id="set" class="displaying">
	<label id="setLabel" >系列:</label> <br>
	<select id="setSelected" name="系列" autofocus>
		<option value="setNotSelected" style="display:none;"></option>
		<%
			Dim sConnection, objConn , objRS ,headerRow, queryStr, hostname
			queryStr=" select set_id, set_name from webone.sets ;"
			sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
			Set objConn = Server.CreateObject("ADODB.Connection") 
			objConn.Open(sConnection) 
			Set objRS = objConn.Execute(queryStr)
			While Not objRS.EOF
		%>
			<option value="<%=objRS.Fields("set_id")%>"><%=objRS.Fields("set_name")%></option>
		<%
			objRS.MoveNext
			Wend
			objRS.Close
			Set objRS = Nothing
			objConn.Close
			Set objConn = Nothing
		%>
	</select>
</div>
&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp
<br>
<div id="subCatLabelDiv">
	<label id="subCatLabel"  class="displaying">分类:</label> 
</div>
<div id = "subCatTxtAreaDiv" class="displaying">
	<textarea id="subCatTxtArea" rows="3" cols="80" onfocus="focusDetection(this)" class="displaying"></textarea> 
</div> <br> 

<button onclick="submitWork()" class="displaying">提交</button>
<script>
function submitWork(){
	var setSelected  				= document.getElementById("setSelected");
	var setIDSelected				= setSelected.options[ setSelected.selectedIndex ].value;
	var adminPageData 				= document.getElementById("adminPageData");
	var checkPointType  			= adminPageData.getAttribute('data_checkpoint_type');
	var subCatTxtArea	 			= document.getElementById("subCatTxtArea").value;
	
	var pViewDelete 				=  "chkType=" 	 + checkPointType + "&" +
									   "setID="		 + setIDSelected;
	var pSaveCheckPoint				=  pViewDelete			 + "&" +
									   "subCat=" + subCatTxtArea;
	var http = new XMLHttpRequest();
	http.open("POST", "saveSubCat.asp", true);
	//Send the proper header information along with the request
	http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	http.onreadystatechange = function(){//Call a function when the state changes.
		if(http.readyState == 4 && http.status == 200){
		   var postSuccess = http.responseText;
		   if (postSuccess == "Y"){
				 if (confirm("提交成功，需要浏览提交结果吗？")){
					navigateTo("viewDeleteSubCat.asp",pViewDelete);				
				  }else {
					 location.reload(true);
				  }
			} else {
				  if (confirm("提交可能完全失败，请联系管理员，仍要浏览提交的结果吗？")){
					 navigateTo("viewDeleteSubCat.asp",pViewDelete);
				  }else {
					 location.reload(true);
				  }
			}
		}
	}
	http.send(pSaveCheckPoint)
}
function chckTypeChanged(chckType){
	var adminPageData 	= document.getElementById("adminPageData");
	if (chckType == "Z"){
       adminPageData.setAttribute('data_checkpoint_type','Z');
	} else if ( chckType == "L"){
	   adminPageData.setAttribute('data_checkpoint_type','L');
	} else if (chckType == "Q"){
	    adminPageData.setAttribute('data_checkpoint_type','Q');
   }
}
function focusDetection(focusedElement) {
var setSelected  	= document.getElementById("setSelected");
var setIDSelected	= setSelected.options[ setSelected.selectedIndex ].value;
console.log("setIDSelected:"+ setIDSelected);
if (focusedElement.id == "subCatTxtArea"  && setIDSelected == "setNotSelected")
  {
	alert("请先选择选择好系列！");
	focusedElement.blur();
   }
}
function navigateTo(url,pList){
		 window.location.href = url + "?" + pList;
}
</script>

</body>
</html>
