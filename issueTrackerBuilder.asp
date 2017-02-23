<!DOCTYPE html>
<html>
<%@ language="VBScript" codepage="65001" %>
<%
Response.CharSet = "utf-8"
session.codepage=65001
%>

<head>

<title>hello..</title>

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

</style>

</head>
<body background="images/singleCatSheet.jpg">
<a href="home.asp">Home</a>

<%
response.expires=-1
Session.Timeout=1440

Dim sConnection, objConn , objRS ,headerRow, queryStr ,checkPointsList,totalCheckPoints,procedureListValues,issueList
checkPointsList = Request.Form("checkPointIdList")

procedureListValues = Request.Form("procedureList")
response.write "<p  id=""procedureListValues"" hidden>" & procedureListValues & "</p>"

issueList = Request.Form("issueList")

if issueList <> "" then
delimiter = Chr(31)
issueListArr = Split(issueList,delimiter)
end if

totalCheckPoints=(len(checkPointsList) - len(replace(checkPointsList, ",", "")) + 1)

queryStr="SELECT CONVERT(sub_cat_id USING utf8) sub_cat_id, set_name , sub_cat_name , checkpoint , fulfill_standard, audit_rule, sub_cat_name,checkpoint_id FROM webone.selfEva where 1= 1  "
queryStr=queryStr & " AND checkpoint_id in " & "(" & checkPointsList & ")"  & " order by  checkpoint_id ASC ;"

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)


'preparing search result buttons
Dim topHead, secondH, buttonSubmit,buttonBack
topHeader=" <tr> <td colspan=""6""> 问题追踪  </td> </tr> "
secondH=" <tr> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>排查发现问题</th>  <th>处置整改措施</th>  <th>整改完成情况</th>  </tr> "

buttonSubmit="<button onclick=""submitProcedures()""> 提交 </button>"
buttonBack="<button onclick=""buttonBack()""> 返回 </button>"

Response.Write "<table id=""selfEvaSheet"" >"
Response.Write topHeader
Response.Write secondH
Dim checkPointIndex	
checkPointIndex=1
While Not objRS.EOF
IF checkPointIndex=1 Then
Response.Write "<tr> <td rowspan=""" & totalCheckPoints & """  >"  & objRS.Fields("set_name")   &  "</td> "
Response.Write "<td rowspan="""      & totalCheckPoints & """ >"  & objRS.Fields("sub_cat_name") &  "</td>  "
Response.Write " <td >" & objRS.Fields("checkpoint") &  "</td> "
Response.Write "<td contenteditable ></td> "
Response.Write "<td contenteditable ></td> "
Response.Write "<td contenteditable ></td>  "

Response.Write "<td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"

End If

IF checkPointIndex>1 Then
Response.Write "<tr> <td >" & objRS.Fields("checkpoint") &  "</td> "
Response.Write "<td contenteditable ></td> "
Response.Write "<td contenteditable ></td> "
Response.Write "<td contenteditable ></td>  "
Response.Write "<td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
End If
checkPointIndex=checkPointIndex+1
objRS.MoveNext
Wend
Response.Write "</table>"
Response.Write "<br><br><br>"
Response.Write buttonBack & buttonSubmit
Response.Write "<input type=""hidden"" name=""orderNumber"" id=""checkPointIndex"" value=""" & ( checkPointIndex - 1 ) & """ />"
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>

<script>
function createArray(length) {
    var arr = new Array(length || 0),
        i = length;
    if (arguments.length > 1) {
        var args = Array.prototype.slice.call(arguments, 1);
        while(i--) arr[length-1 - i] = createArray.apply(this, args);
    }
    return arr;
}

function printArrays(arrays ) {
	for(var i = 0; i < arrays.length; i++) {
			var array = arrays[i];
			for(var j = 0; j < array.length; j++) {
				alert( '[' + i + ']' + '[' + j + '] is '  + arrays[i][j] );
			}
		}	
}

function getCellValues(arrays,table) {
	for (var r = 0, n = table.rows.length; r < n; r++) {
		for (var c = 0, m = table.rows[r].cells.length; c < m; c++) {
		  if (  r > 1 ) {
			if ( r == 2 ){
				 if (c > 2){
					 arrays[ r - 2 ][ c - 3 ] = table.rows[r].cells[c].innerHTML;
				 }						  
			 } else 
			 {  // r > 2
				if ( c > 0 ){
					arrays[ r - 2 ][ c - 1 ] = table.rows[r].cells[c].innerHTML;
				}
				 //alert('row number ' + r  + 'cell number ' + c + ' ends'); 
			 }
		   }  //if (  r > 1 ) 
		}
	}
}
	
function generateIssueList(arr){
	var fieldsList="nothingToBeTold"
	var fieldDimitor=String.fromCharCode(31);
	for(var i = 0; i < arr.length; i++) {
			var array = arr[i];
			for(var j = 0; j < array.length; j++) {
				  if (fieldsList == 'nothingToBeTold'){
					 fieldsList = arr[i][j];
				  }else{
					 fieldsList = fieldsList + fieldDimitor + arr[i][j];
				  }
			}
		}
	return fieldsList
}

function navigates(navigateTo){
	window.location.href = navigateTo
}; 

function post(url, params, method,navigateToUrl,newUrl){
	method = method || "post"; // Set method to post by default if not specified.
	navigateToUrl = navigateToUrl || "N"; // Set method to post by default if not specified.
	var proceduresAndIssues = ''
	var elementCount =1
	for(var key in params) {
		if(params.hasOwnProperty(key)) {
		   if (elementCount == 1) {
			  proceduresAndIssues =  proceduresAndIssues +  key + '=' +  window.encodeURIComponent(params[key]);
		   } else {
			  proceduresAndIssues =  proceduresAndIssues +  '&' + key + '=' +  window.encodeURIComponent(params[key]);
		   }
		   elementCount = elementCount + 1;
		 }
	}
	if ( navigateToUrl == "N" ) { //Not navigate to the  url, so navigate to newUrl
		var http = new XMLHttpRequest();
		http.open("POST", url, true);
		//Send the proper header information along with the request
		http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		http.onreadystatechange = function(){//Call a function when the state changes.
			if(http.readyState == 4 && http.status == 200){
			   alert(http.responseText); //debug 
			   navigates(newUrl)
			 }
		   }
		http.send(proceduresAndIssues)
	} else { 
    // The rest of this code assumes you are not using a library.
    // It can be made less wordy if you use one.
		var form = document.createElement("form");
		form.setAttribute("method", method);
		form.setAttribute("action", url);
		for(var key in params) {
			if(params.hasOwnProperty(key)) {
				var hiddenField = document.createElement("input");
				hiddenField.setAttribute("type", "hidden");
				hiddenField.setAttribute("name", key);
				hiddenField.setAttribute("value", params[key]);
				  console.log('this is params (' + key + ') for sure. Value: ' + params[key]); 
				form.appendChild(hiddenField);
			 }
		}
		document.body.appendChild(form);
		form.submit();
	}
}

function submitProcedures(){
	var table = document.getElementById('selfEvaSheet');
	var checkPointIndex = document.getElementById('checkPointIndex').value;
	var issuesTracked = createArray(checkPointIndex,4); 
	getCellValues(issuesTracked,table);
	var procedureListValues = document.getElementById('procedureListValues').innerHTML;
	var parameters = {
	  issueList: generateIssueList(issuesTracked),
	  procedureList: procedureListValues
	}
	/*
	for (var name in parameters) {
	  if (parameters.hasOwnProperty(name)) {
		console.log('this is fog (' + name + ') for sure. Value: ' + parameters[name]);
	  }
	  else {
		console.log(name); // toString or something else
	  }
	}
	*/
	post("saveInDB.asp",parameters,"post","N","http://localhost:88/selfEvaNavigation.asp?navType=selfEva")
}
function buttonBack(){
	var table = document.getElementById('selfEvaSheet');
	var checkPointIndex = document.getElementById('checkPointIndex').value;
	var issuesTracked = createArray(checkPointIndex,4); 
	getCellValues(issuesTracked,table);
    var procedureListValues = document.getElementById('procedureListValues').innerHTML;
	var parameters = {
	  issueList: generateIssueList(issuesTracked),
	  procedureList: procedureListValues
	}
	post(document.referrer,parameters,"post","Y","http://localhost:88/selfEvaNavigation.asp?navType=selfEva")

}
</script>

</body>
</html>