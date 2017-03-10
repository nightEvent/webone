<!DOCTYPE html>
<html>
<%@ language="VBScript" codepage="65001" %>
<%
Response.CharSet = "utf-8"
session.codepage=65001
%>

<head>
<link rel="icon" href="images/favicon.ico" type="image/x-icon">
<title>浏览系列</title>


<style>
	hr { 
		display: block;
		margin-top: 0.1em;
		margin-bottom: 0.05em;
		margin-left: auto;
		margin-right: auto;
		border-style: inset;
		border-width: 1px;
	}
	
  
.selectStyle {

	
	border: #FEFEFE 1px solid;
	-webkit-border-radius: 3px;
	border-radius: 3px;
	-webkit-box-shadow: inset 0px 0px 10px 1px #FEFEFE;
	box-shadow: inset 0px 0px 10px 1px #FEFEFE;
   }


.selectStyle select {
   background: transparent;
   width: 170px;
   font-size:7pt;
   color:grey;
   border: 0;
   border-radius: 0;
   height: 28px;
   -webkit-appearance: none;
   
   }
.selectStyle select:focus {
    outline: none;

}
</style>

</head>

<body >
<a href="#" onclick="homeClicked()">Home</a> <br>
<%
response.expires=-1
Session.Timeout=1440
Dim setID,chkType,sConnection, objConn , objRS ,headerRow, queryStr,firstRecordFlag,chkTypeName
setID 			= Request.querystring("setID")
chkType			= Request.querystring("chkType")

queryStr="SELECT sub_cat_id, sub_cat_name ,chk_Type,set_name,set_id FROM webone.subCatNav where 1= 1  "
queryStr=queryStr & " AND chk_Type = '" & chkType & "' "
queryStr=queryStr & " AND set_id  = " & setID & " order by  sub_cat_id ASC ;"

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)
if chkType = "Q" then
	chkTypeName="全面排查"
elseif chkType = "L" then
	chkTypeName="临时排查"
elseif chkType = "Z" then
	chkTypeName="自查"
else 
	chkTypeName="未知类型排查"
end if

firstRecordFlag="Y"
While Not objRS.EOF
if firstRecordFlag="Y" then
	response.write "<h3>系列：" & objRS.Fields("set_name")  & "<br>  类型：" & chkTypeName & "</h3>"
	response.write "<hr>"
	response.write "<select class=""selectStyle"" id=""subCat"" multiple>"
	response.write "<option value=""" & objRS.Fields("sub_cat_id") & """>" & objRS.Fields("sub_cat_name") & "</option>"
	firstRecordFlag="N"
else
response.write "<option value=""" & objRS.Fields("sub_cat_id") & """>" & objRS.Fields("sub_cat_name") & "</option>"
end if
objRS.MoveNext
Wend
response.write "</select>"
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>

<br>
<br>
<br>
<button onclick=buttonBack()> 返回添加风险点页面 </button> &nbsp &nbsp
<button onclick="deleteSelected()"> 删除选定的风险点</button>

<script>
function deleteSelected(){
    var select= document.getElementById("subCat");
	var subCatList=arrayToJoin(getSelectValues(select),44);
	var parameters = "subCatList=" + subCatList + "&deleteType=S";
	console.log(parameters)
	var http = new XMLHttpRequest();
	http.open("POST", "deleteTheAdded.asp", true);
	http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	http.onreadystatechange = function(){
	if(http.readyState == 4 && http.status == 200){
		   alert(http.responseText);
		   location.reload(true);
		 }
	}
	http.send(parameters)
}

function arrayToJoin(arr,separatorASCII){
var fieldDimitor=String.fromCharCode(separatorASCII)
return arr.join(fieldDimitor);
}

function arrayToSeparatedList(arr,separatorASCII){
	var SeparatedList="nothingToBeTold";
	var fieldDimitor=String.fromCharCode(separatorASCII); //31 for unit separator, 44 for comma
	for(var i = 0; i < arr.length; j++) {
		if (SeparatedList == 'nothingToBeTold'){
		 SeparatedList = arr[i];
		}else{
		 SeparatedList = SeparatedList + fieldDimitor + arr[i];
		}
	}
	return SeparatedList;
}

function getSelectValues(select) {
  var result = [];
  var options = select && select.options;
  var option;

  for (var i=0, iLen=options.length; i<iLen; i++) {
    option = options[i];
    if (option.selected) {
      result.push(option.value);
    }
  }
  return result;
}
function buttonBack(){
//var currentPageData 	= document.getElementById("currentPageData");
//var checkPointType  	= currentPageData.getAttribute('data_checkpoint_type');
navigates("addSubCat.asp")
}
function homeClicked(){
navigates("home.asp")
}
function navigates(navigateTo){
 window.location.href = navigateTo
};
</script>

</body>
</html>