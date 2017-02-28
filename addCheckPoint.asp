<!DOCTYPE html>
<html>
<head>
    <title>太平人寿风险排查系统</title>
	<meta charset="utf-8" />
	<!-- <link rel="stylesheet" href="/w3.css"> -->
<style>

a.homeLink:hover,a.homeLink.active{
  text-decoration: underline;
}
</style>
<!--JQuery source used-->
<script src="lib/jquery.min.js"></script>
</head>
<a class="homeLink" href="http://localhost:88/home.asp">Home</a> <br> <br>
<body background="images/homeBackground.jpg" >
<%
Dim checkPointType
checkPointType = Request.QueryString("checkPointType")
response.write "<div id=""adminPageData"" data_checkpoint_type=""" & checkPointType & """ auditRule-changed=""N""></div>"
%>
<label>系列</label> <br>

<select id="setSelected" name="系列" onchange="populateSubCat()" autofocus>
<option value="XX" style="display:none;"></option>
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

&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp
<br>
<br>
<lable>分类</label> &nbsp 
<div id = "subCatDiv">
<select id="subCatSelected">
<option></option>
</select>
</div> <br><br>

<lable>排查方法：</label> &nbsp 
<div id = "auditRuleDiv">
<textarea id="auditRule" rows="10" cols="60" onfocus="focusDetection(this)"></textarea>
</div>
<br>
<form action="saveCheckPoints.asp" method="post" id="checkPointForm">
  详细风险点： <br> <textarea id="checkPointContent" name="checkPointName" onfocus="focusDetection(this)" rows="10" cols="60" form="checkPointForm">请在这个输入框中添加详细风险点...</textarea>
  <br> 
  合规要求：<br> <textarea id="checkPointFulfillStandard" name="fulfillStandard" onfocus="focusDetection(this)" rows="10" cols="60" form="checkPointForm" >请在这个输入框中添加合规要求...</textarea>
  <br>
</form>
<button onclick="submitWork()">提交</button>
<script>
function submitWork(){
//checkpointType,subCatSelected,auditRule
var adminPageData 				= document.getElementById("adminPageData");
var checkPointType  			= adminPageData.getAttribute('data_checkpoint_type');
var subCatSelected 				= document.getElementById("subCatSelected");
var subCatIDSelected  			= subCatSelected.options[ subCatSelected.selectedIndex ].value;
var auditRuleChanged			= adminPageData.getAttribute('auditRule-changed');
if (auditRuleChanged == 'Y') {
	var auditRuleEitherNewOrUpdated = document.getElementById("auditRule").value;
}
var checkPointContent			= document.getElementById("checkPointContent").value;
var checkPointFulfillStandard	= document.getElementById("checkPointFulfillStandard").value;
var pViewDelete 				=  "checkPointType=" 			  + checkPointType 				+ "&" +
								   "subCatID="					  + subCatIDSelected;
var pSaveCheckPoint				=  pViewDelete		    	 	  								+ "&" +
								   "auditRuleEitherNewOrUpdated=" + auditRuleEitherNewOrUpdated + "&" + 
								   "auditRuleChanged="			  + auditRuleChanged			+ "&" +
								   "checkPointContent=" 		  + checkPointContent 			+ "&" +
								   "checkPointFulfillStandard="   + checkPointFulfillStandard;
var http = new XMLHttpRequest();
http.open("POST", "saveCheckPoint.asp", true);
//Send the proper header information along with the request
http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
http.onreadystatechange = function(){//Call a function when the state changes.
	if(http.readyState == 4 && http.status == 200){
	   var postSuccess = http.responseText;
	   if (postSuccess == "Y"){
			 if (confirm("提交成功，需要浏览提交结果吗？")){
				navigateTo("viewDeleteCheckPoint.asp",pViewDelete);				
			  }else {
				 location.reload(true);
			  }
		} else {
			  if (confirm("提交可能完全失败，请联系管理员，还需要浏览提交的结果吗？")){
				 navigateTo("viewDeleteCheckPoint.asp",pViewDelete);
			  }else {
				 location.reload(true);
			  }
		}
	}
}
http.send(pSaveCheckPoint)
}
function auditRuleChanged(){
//set audit rule change flag to Y, so that we later know we will need to update/insert audit rule
	var adminPageData 	= document.getElementById("adminPageData");
	adminPageData.setAttribute('auditRule-changed','Y');
}
function loadAuditRule(){
	var url 			= "checkPointReturns.asp";
	var subCatSelected  = document.getElementById("subCatSelected");
	var subCatID		= subCatSelected.options[ subCatSelected.selectedIndex ].value;
	//you found it's already selected then your responsibility to store it in global pool for sumbmission reference
	var adminPageData 	= document.getElementById("adminPageData");
	
	var checkPointType  = adminPageData.getAttribute('data_checkpoint_type');
	var Params 			= "RequestType=" + window.encodeURIComponent("AR") + "&" +
						  "checkPointType=" + window.encodeURIComponent(checkPointType) +  "&" +
						  "subCatID=" +  window.encodeURIComponent(subCatID);
	var http 			= new XMLHttpRequest();
	http.open("POST", url, true);
	//Send the proper header information along with the request
	http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	http.onreadystatechange = function(){//Call a function when the state changes.
		if(http.readyState == 4 && http.status == 200){
		   var auditRuleLoaded  = http.responseText;
		   //alert('auditRuleLoaded is: '+ auditRuleLoaded);
		   	var auditRule = document.getElementById("auditRule");
			auditRule.remove();
			var auditRuleNew = document.createElement("textarea");
			auditRuleNew.id = "auditRule";
			auditRuleNew.onchange = function(){auditRuleChanged()};
			auditRuleNew.onfocus  = function(){focusDetection(this)};
			auditRuleNew.rows= "10";
			auditRuleNew.cols= "60";
			var auditRuleDiv = document.getElementById("auditRuleDiv");
			auditRuleDiv.appendChild(auditRuleNew);
			var node = document.createTextNode(auditRuleLoaded);
			auditRuleNew.appendChild(node);
		 }
	   }
	http.send(Params);
}
function populateSubCat(){
	var url 			= "checkPointReturns.asp";
	var setSelected 	= document.getElementById("setSelected");
	var setID           = setSelected.options[ setSelected.selectedIndex ].value;
	var adminPageData 	= document.getElementById("adminPageData");
	var checkPointType  = adminPageData.getAttribute('data_checkpoint_type');
	var Params 			= 'RequestType=' + 'SC' + '&' +
						  'checkPointType=' + checkPointType  + '&' +
						  'setID='  + setID;
	var http = new XMLHttpRequest();
	http.open("POST", url, true);
	//Send the proper header information along with the request
	http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	http.onreadystatechange = function(){//Call a function when the state changes.
		if(http.readyState == 4 && http.status == 200){
			var subCatResponseTxt  = http.responseText;
			var subCatSelect = document.getElementById("subCatSelected");
			subCatSelect.remove();
			var subCatSelectNew = document.createElement("select");
			subCatSelectNew.id = "subCatSelected";
			subCatSelectNew.onchange = function(){loadAuditRule()};
			var option = document.createElement("option");
			option.value = "";
			option.text = "";
			option.style = "display:none;";
			subCatSelectNew.appendChild(option);
			var subCatDiv = document.getElementById("subCatDiv");
			subCatDiv.appendChild(subCatSelectNew);
			var delimiter=String.fromCharCode(31);
			var array = subCatResponseTxt.split(delimiter);
			var catIdNdName = createArray(array.length/2,2);
			for (i = 0 ; i < array.length; i++){
			   if( i%2 == 0 ){
				catIdNdName[Math.floor(i/2)][0] = array[i];
			   } else {
				catIdNdName[Math.floor((i-1)/2)][1] = array[i];
			   }
			}
			//Create and append the options
			for (var i = 0; i < catIdNdName.length; i++) {
				var option = document.createElement("option");
				option.value = catIdNdName[i][0];
				option.text = catIdNdName[i][1];
				subCatSelectNew.appendChild(option);
			}
		 }
	   }
	http.send(Params)
}
function createArray(length) {
    var arr = new Array(length || 0),
        i = length;
    if (arguments.length > 1) {
        var args = Array.prototype.slice.call(arguments, 1);
        while(i--) arr[length-1 - i] = createArray.apply(this, args);
    }
    return arr;
}

function focusDetection(focusedElement) {
var subCatSelected  = document.getElementById("subCatSelected");
var subCatID		= subCatSelected.options[ subCatSelected.selectedIndex ].value;
console.log("subCatID:"+ subCatID);
console.log("focusedElement.id:"+ focusedElement.id);
var adminPageData	= document.getElementById("adminPageData");
var checkPointType	= adminPageData.getAttribute('data_checkpoint_type');
if ((focusedElement.id == "auditRule" || focusedElement.id == "checkPointContent"  || focusedElement.id == "checkPointFulfillStandard") 
    && subCatID == "")
  {
	alert("请先选择选择好分类！");
	focusedElement.blur();
   }
}
function navigateTo(url,pList){
		 window.location.href = url + "?" + pList;
	}
</script>

</body>
</html>
