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

<div id="adminPageData" data_checkpoint_type="Z" auditRule-changed="N"></div>

<br>
<div id="set" class="displaying">
	<label id="setLabel" >系列:</label> <br>
	<select id="setSelected" name="系列" onchange="populateSubCat()" autofocus>
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
<label id="subCatLabel" class="displaying">分类:</label> 
<div id = "subCatDiv" class="displaying">
	<select id="subCatSelected">
	<option></option>
	</select>
</div> <br>
<div id="auditRuleLabelDiv">
	<label id="auditRuleLabel"  class="displaying">排查方法：</label> 
</div>
<div id = "auditRuleDiv" class="displaying">
	<textarea id="auditRuleTxtArea" rows="10" cols="60" onfocus="focusDetection(this)" class="displaying"></textarea> 
</div> <br> 
<div id="checkPointDiv" class="displaying">
	详细风险点： <br> <textarea id="checkPointContent" name="checkPointName" onfocus="focusDetection(this)" rows="10" cols="60" form="checkPointForm">请在这个输入框中添加详细风险点...</textarea>
	<br> 
	合规要求：<br> <textarea id="checkPointFulfillStandard" name="fulfillStandard" onfocus="focusDetection(this)" rows="10" cols="60" form="checkPointForm" >请在这个输入框中添加合规要求...</textarea>
	<br>
</div>
<button onclick="submitWork()" class="displaying">提交</button>
<script>
function chckTypeChanged(chckType){
	if (chckType == "Z"){
       location.reload(true);
	} else if ( chckType == "L"){
	   //location.reload(true);
	   //var radioClicked = document.getElementById('chckTypeL');
	   var adminPageData 	= document.getElementById("adminPageData");
	   adminPageData.setAttribute('data_checkpoint_type','L');
	   populateSubCat();
	   if (!document.getElementById('auditRuleLabel')){
			var auditRuleLabel				= document.createElement("LABEL");
				auditRuleLabel.id 			= "auditRuleLabel";
				auditRuleLabel.className	= "displaying";
				auditRuleLabel.innerHTML	= "排查方法：";
			var auditRuleLabelDiv 			= document.getElementById("auditRuleLabelDiv");
				auditRuleLabelDiv.appendChild(auditRuleLabel);
		}
		if (!document.getElementById('auditRuleTxtArea')){
			var auditRuleTxtArea 			= document.createElement("textarea");
				auditRuleTxtArea.id 		= "auditRuleTxtArea";
				auditRuleTxtArea.rows 		= "10";
				auditRuleTxtArea.cols 		= "60";
			  //auditRuleTxtArea.onfocus	= "focusDetection(this)";
				auditRuleTxtArea.onfocus	= function(){focusDetection(this)};
				auditRuleTxtArea.onchange 	= function(){auditRuleChanged()};
				auditRuleTxtArea.className	= "displaying";
			var auditRuleDiv 				= document.getElementById("auditRuleDiv");
				auditRuleDiv.appendChild(auditRuleTxtArea);
		}
	} else if (chckType == "Q"){
		var adminPageData 	= document.getElementById("adminPageData");
	    adminPageData.setAttribute('data_checkpoint_type','Q');
		populateSubCat();
		document.getElementById("setLabel").className 		= "displaying";
		document.getElementById("subCatLabel").className 	= "displaying";
		document.getElementById("setSelected").className 	= "displaying";
		document.getElementById("subCatDiv").className 		= "displaying";
		document.getElementById("checkPointDiv").className 	= "displaying";
		rmAuditRule("auditRuleLabel","auditRuleTxtArea");
		function rmAuditRule(auditRuleLabel,auditRuleTxtArea){
			if (document.getElementById(auditRuleLabel)){
				var auditRule = document.getElementById(auditRuleLabel);
					auditRule.remove();
			}
			if(document.getElementById(auditRuleTxtArea)){
				var auditRulex = document.getElementById(auditRuleTxtArea);
					auditRulex.remove();
			}
	   }
   }
}

function submitWork(){
	//checkpointType,subCatSelected,auditRule
	var adminPageData 				= document.getElementById("adminPageData");
	var checkPointType  			= adminPageData.getAttribute('data_checkpoint_type');
	var subCatSelected 				= document.getElementById("subCatSelected");
	var subCatIDSelected  			= subCatSelected.options[ subCatSelected.selectedIndex ].value;
	var auditRuleChanged			= adminPageData.getAttribute('auditRule-changed');
	var auditRuleEitherNewOrUpdated = "no need for insert or update";
	if (auditRuleChanged == 'Y') { //only Y, will then do update or insert in saveCheckPoint.asp
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
		    var auditRuleLoaded  	= http.responseText;
		   	var auditRuleTxtArea 	= document.getElementById("auditRuleTxtArea");
			auditRuleTxtArea.remove();
			var auditRuleNew 		= document.createElement("textarea");
			auditRuleNew.id 		= "auditRuleTxtArea";
			auditRuleNew.onchange 	= function(){auditRuleChanged()};
			auditRuleNew.onfocus  	= function(){focusDetection(this)};
			auditRuleNew.rows		= "10";
			auditRuleNew.cols		= "60";
			var node = document.createTextNode(auditRuleLoaded);
			auditRuleNew.appendChild(node);
			var auditRuleDiv = document.getElementById("auditRuleDiv");
			auditRuleDiv.appendChild(auditRuleNew);
		 }
	   }
	http.send(Params);
}
function populateSubCat(){
	var url 			= "checkPointReturns.asp";
	var setSelected 	= document.getElementById("setSelected");
	var setID           = setSelected.options[ setSelected.selectedIndex ].value;
	if ( setID == "setNotSelected") {
	   return;
	}
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
			var option = document.createElement("option");
			option.value = "";
			option.text = "";
			option.style = "display:none;";
			subCatSelectNew.appendChild(option);
			var subCatDiv = document.getElementById("subCatDiv");
			subCatDiv.appendChild(subCatSelectNew);
			if (document.getElementById("chckTypeL").checked || document.getElementById("chckTypeZ").checked  ) {
				subCatSelectNew.onchange = function(){loadAuditRule()};
			}
			if (subCatResponseTxt == "NoSubCatReturn"){
			    alert("没有相应的系列返回，请检查是否添加过！");
			    return;
			}
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
if ((focusedElement.id == "auditRuleTxtArea" || focusedElement.id == "checkPointContent"  || focusedElement.id == "checkPointFulfillStandard") 
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
