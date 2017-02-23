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
//		td {vertical-align:center;text-align:justify;}
		tr {text-align:justify}
</style>

</head>

<body background="images/singleCatSheet.jpg">
<a href="#" onclick="homeClicked()">Home</a>
<%
response.expires=-1
Session.Timeout=1440
Dim subCat,checkPointType,sConnection, objConn , objRS ,headerRow, queryStr 
subCat 			= Request.querystring("subCatID")
checkPointType	= Request.querystring("checkPointType")
queryStr="SELECT CONVERT(sub_cat_id USING utf8) sub_cat_id, set_name , sub_cat_name , checkpoint , fulfill_standard, audit_rule,checkpoint_id FROM webone.selfEva where 1= 1  "
queryStr=queryStr & " AND checkpoint_type = '" & checkPointType & "' "
queryStr=queryStr & " AND sub_cat_id  = " & "'" & subCat & "'"  & " order by  checkpoint_id ASC ;"

sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 

Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 


'getting record user previously input
Dim procedureList, arr(),arrCB(),ind,cnt,cntCB,loadInputHistory
procedureList=Request.Form("procedureList")

delimiter = Chr(31)
ind=0
cnt=0
cntCB=0
loadInputHistory="N"
ReDim Preserve arr(10)
ReDim Preserve arrCB(10)
previousPage = Request.ServerVariables("HTTP_REFERER")
if procedureList <> "" AND previousPage <> "http://localhost:88/selfEvaNavigation.asp?navType=selfEva" Then  'checkpointsNavigation.asp
loadInputHistory="Y"
arrays=Split(procedureList,delimiter)
	for each x in arrays
		  if (ind = 0) OR ((ind mod 3) = 0) then 
				arr(cnt)=x
				cnt = cnt + 1
				if (cnt > 2) then
				 ReDim Preserve arr(UBound(arr) + 1)
				end if
		  elseif ((ind mod 3) = 1) then
				if ( x = "") OR ( x = "N" ) then
				  arrCB(cntCB)="unchecked"
				else
				  arrCB(cntCB)="checked"
				end if
				cntCB = cntCB + 1
				if (cntCB > 2) then
				 ReDim Preserve arrCB(UBound(arrCB) + 1)
				end if
		  end if
		  ind= ind + 1
	next
End If


'preparing search result buttons
Dim topHead, secondH, sqlCount,count, buttonSubmit,buttonBack
topHeader=" <tr> <td colspan=""7""> 当前已添加在分类的项目  </td> </tr> "
secondH=" <tr> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>合规要求</th>  <th>排查方法</th> <th style=""display:none;"" > 排查过程</th>  <th>勾选要删除的项</th>  </tr> "
sqlCount="select count(1) as checkPointCnt from webone.checkpoints where checkpoint_type = '" & checkPointType & "' and sub_cat_id = " & subCat & " ;"
Set objRS = objConn.Execute(sqlCount)
while Not objRs.EOF
	count=objRs.fields("checkPointCnt")
	objRS.MoveNext
wend
buttonSubmit="<button onclick=""startChecking()""> 删除选定的风险点 </button>"
buttonBack="<button onclick=""buttonBack()""> 返回添加风险点页面 </button>"
Response.Write "<table id=""selfEvaSheet"" >"
Response.Write topHeader
Response.Write secondH
Dim checkPointsCount
checkPointsCount=1
Set objRS = objConn.Execute(queryStr)
While Not objRS.EOF
IF checkPointsCount=1 Then
	Response.Write "<tr > <td rowspan=""" & count & """  >"  & objRS.Fields("set_name")      &  "</td> "
	Response.Write "<td rowspan="""      & count & """  >"  & objRS.Fields("sub_cat_name") &  "</td>  "
	Response.Write "  <td>" &  " <a href=""#"" onmouseenter=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")"" >"  &  objRS.Fields("checkpoint") &  "</a> "  &   "</td>  <td   > " & objRS.Fields("fulfill_standard") & "</td> "
	Response.Write "<td rowspan=""" & count & """ >"  & objRS.Fields("audit_rule") & "</td>"
	Response.Write "<td  style=""display:none;""> 排查过程 </td>"
	Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" unchecked >  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
End If
IF checkPointsCount>1 Then
	Response.Write "<tr> <td >" &  " <a href=""#"" onmouseenter=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")"" >"  &  objRS.Fields("checkpoint") &  "</a> "  &  "</td>  <td   > " & objRS.Fields("fulfill_standard") & "</td> "
	Response.Write "<td  style=""display:none;""> 排查过程 </td>"
	Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" unchecked >  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
End If
checkPointsCount=checkPointsCount+1
objRS.MoveNext
Wend
Response.Write "</table>"
Response.Write "<br>"
Response.Write buttonBack
Response.Write " "
Response.Write buttonSubmit
Response.Write "<input type=""hidden"" name=""orderNumber"" id=""checkPointsCount"" value=""" & ( checkPointsCount - 1 ) & """ />"
Response.Write "<br>"
Response.Write "<p id=""linkedRules""></p>"
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>

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

function getCellValues(audit_procedures,table) {
        //alert('created');	
        for (var r = 0, n = table.rows.length; r < n; r++) {
            for (var c = 0, m = table.rows[r].cells.length; c < m; c++) {
			  if (  r > 1 ) {
			    if ( r == 2 ) {
				     //alert('when row is  2 ' ); 
					if ( c > 4 ) { 
					   // alert('row 2' + ', cell ' + c ); 
						if ( c == 6 ) // checkbox at cell number 6
						{
							var idddd = (r - 1 )*1000;
							 if ( document.getElementById(idddd).checked  )  { 
								audit_procedures[ r - 2 ][ c - 5 ] = "Y";
								//alert('checked  captured and what is stored in ' + '[' + (r-2) + ']['+ (c-5)+ '] is ' + audit_procedures[ r - 2 ][ c - 5 ] );
							 }
							 else {
								audit_procedures[ r - 2 ][ c - 5 ] = "N";
								//alert('unchecked captured and what is stored in '  + '[' + (r-2) + ']['+ (c-5)+ '] is ' + audit_procedures[ r - 2 ][ c - 5 ] );
							}
						 } else   // if ( c <> 6 ) , cell number 3, 5, 6, 7 will be procedure, sub_cat_id , checkpoint_id
						 {
					          //	   alert('value is ' + table.rows[r].cells[c].innerHTML ); 
						   audit_procedures[ r - 2 ][ c - 5 ] = table.rows[r].cells[c].innerHTML;
						   //alert(table.rows[r].cells[c].innerHTML + 'captured and what is stored in ' + '[' + (r-2) + ']['+ (c-5)+ '] is '  + audit_procedures[ r - 2 ][ c - 5 ] );
						 }
					 }  //if ( c > 1 )
					  //alert('row 2 , cell number ' + c + ' ends');
                 } else 
				 {   // r > 2
				 // alert('row number ' + r + ', cell number' + c );
					if ( c > 1 ) { 
						if ( c == 3 ) {  // checkbox at cell number 3
							var idddd = (r - 1 )*1000; 
							 if ( document.getElementById(idddd).checked  )  { 
								audit_procedures[ r - 2 ][ c - 2 ] = "Y";
								//alert('checked captured');
							 }
							 else {
								audit_procedures[ r - 2 ][ c - 2 ] = "N";
								//alert('unchecked captured');
							}
						 } 
							else
							{
							   audit_procedures[ r - 2 ][ c - 2 ] = table.rows[r].cells[c].innerHTML;
							   //alert(table.rows[r].cells[c].innerHTML+' captured');
							}
					   }
					 //alert('row number ' + r  + 'cell number ' + c + ' ends'); 
				 }
			   }  //if (  r > 1 ) 
            }
        }

    }

function procedureListBuilder(arrays){
	var http = new XMLHttpRequest();
	var url = "saveProcedureInSession.asp";
	var params;
	//procedure, flag ,sub_cat_id ,checkpoint_id
	//var params = "insert into webone.audit_procedure (procedure,issue_found_flag,checkpoint_id) ";
	var procedureList="nothingToBeTold"
	var fieldDimitor=String.fromCharCode(31);
	for(var i = 0; i < arrays.length; i++) {
			var array = arrays[i];
			for(var j = 0; j < array.length; j++) {
			   if ( j !== 2) {
				  if (procedureList == 'nothingToBeTold'){
					 procedureList = arrays[i][j];
				  }else{
					 procedureList = procedureList + fieldDimitor + arrays[i][j];
				  }
				}
			}
		}
	//alert(procedureList);
	return procedureList
}

function navigateToIssueTracker(inlist){
	var navigateTo="http://localhost:88/issueTrackerBuilder.asp?checkPointIdList=" + inlist;
	//alert('navigate to ' + navigateTo);
	//window.location=navigateTo;
	 window.location.href = navigateTo
}

function getInlist(arrays ) {
	var inlist = 'NULL';
	for(var i = 0; i < arrays.length; i++)
	  {
	   if ( arrays[i][1] == 'Y' ) {
			if ( i > 0 ) 
			{
			  if ( inlist !== 'NULL' ) {
				inlist = inlist + ',';
			  }          	  
			}
			if (inlist == 'NULL'){
			  inlist = arrays[i][3];
			}else
			{
			inlist = inlist + arrays[i][3];
			}
		  }
	   }    

	return inlist
}

function startChecking(){
	var table = document.getElementById('selfEvaSheet');
	var checkPointsCount = document.getElementById('checkPointsCount').value;
	var audit_procedures = createArray(checkPointsCount, 4 ); 
	getCellValues(audit_procedures,table);
	//printArrays(audit_procedures);
	var parameters = "checkPointIdList=" + getInlist(audit_procedures);
	console.log(parameters)
	var http = new XMLHttpRequest();
	http.open("POST", "deleteCheckPoints.asp", true);
	http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	http.onreadystatechange = function(){
	if(http.readyState == 4 && http.status == 200){
		   alert(http.responseText);
		   location.reload(true);
		   //navigates(document.referrer)
		 }
	   }
	http.send(parameters)
}

function navigates(navigateTo){
 window.location.href = navigateTo
};


function homeClicked(){
navigates("home.asp")
}

function buttonBack(){
navigates("addCheckPoint.asp")
}
</script>

</body>
</html>