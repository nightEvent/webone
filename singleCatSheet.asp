<!DOCTYPE html>
<html>
<%@ language="VBScript" codepage="65001" %>
<%
Response.CharSet = "utf-8"
session.codepage=65001
%>

<head>
<title>风险点排查</title>
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
//		td {vertical-align:center;text-align:justify;}
		tr {text-align:justify}
</style>

</head>

<body background="images/singleCatSheet.jpg">
<a href="#" onclick="homeClicked()">Home</a>
<%
response.expires=-1
Session.Timeout=1440
Dim subCatId,sConnection, objConn , objRS ,headerRow, queryStr 
subCatId = Request.querystring("subCatId")

queryStr="SELECT sub_cat_id,chk_Type, set_name , account, sub_cat_name , checkpoint , fulfill_standard, audit_rule,checkpoint_id FROM webone.selfEva where 1= 1  "
queryStr=queryStr & " AND sub_cat_id  = " & "'" & subCatId & "'"  & " order by  checkpoint_id ASC ;"
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3"
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 

'getting record user previously input
Dim procedureList, arr(),arrCB(),ind,cnt,cntCB,loadInputHistory
procedureList=Request.Form("procedureList")
issueList=Request.Form("issueList")

response.write "<p  id=""issueList"" hidden>" & issueList & "</p>"

delimiter = Chr(31)
ind=0
cnt=0
cntCB=0
loadInputHistory="N"
ReDim Preserve arr(10)
ReDim Preserve arrCB(10)
previousPage = Request.ServerVariables("HTTP_REFERER")
if procedureList <> "" then 'AND previousPage <> "http://localhost:88/selfEvaNavigation.asp?navType=selfEva" Then  'checkpointsNavigation.asp  'seems not required
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
Dim topHead, secondH,sqlCount, count,startCheck, buttonSubmit,buttonBack
sqlCount="select count(1) as checkPointCnt from webone.checkpoints where checkpoint_type = '" & checkPointType & "' and sub_cat_id = " & subCatId & " ;"
Set objRS = objConn.Execute(sqlCount)
while Not objRs.EOF
	count=objRs.fields("checkPointCnt")
	objRS.MoveNext
wend
topHeader=" <tr> <td colspan=""7""> 甘肃分公司内控风险自查表  </td> </tr> "
secondH=" <tr> <th>系列</th> <th>分类</th> <th>详细风险点</th>  <th>合规要求</th>  <th>排查方法</th>  <th>排查经过</th> <th>发现问题</th>  </tr> "
startCheck="<button onclick=""startChecking()""> 排查开始 </button>"
buttonSubmit="<button onclick=""startChecking()""> 进入问题追踪页面 </button>"
buttonBack="<button onclick=""buttonBack()""> 返回系列 </button>"
Response.Write "<table id=""selfEvaSheet"" >"
Response.Write topHeader
Response.Write secondH
Dim checkPointsCount
checkPointsCount=1
Set objRS = objConn.Execute(queryStr)
While Not objRS.EOF
IF checkPointsCount=1 Then
	Response.Write "<div id=""singleCatSheetPgData"" chkType=""" & objRS.Fields("chk_Type") & """ account=""" & objRS.Fields("account") &  """> </div>"
	Response.Write "<tr > <td rowspan=""" & count & """  >"  & objRS.Fields("set_name")      &  "</td> "
	Response.Write "<td rowspan="""      & count & """  >"  & objRS.Fields("sub_cat_name") &  "</td>  "

	Response.Write "  <td   >" &  " <a href=""#"" onmouseenter=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")"" >"  &  objRS.Fields("checkpoint") &  "</a> "  &   "</td>  <td   > " & objRS.Fields("fulfill_standard") & "</td> "

	Response.Write "<td rowspan=""" & count & """ >"  & objRS.Fields("audit_rule") & "</td>"
	IF loadInputHistory = "Y" THEN
		Response.Write "<td contenteditable            onclick=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")""         >" & arr(checkPointsCount-1) & "</td> " 'when row number is 2, read from cell NUMBER 5 through 8, when cell number is 6 , it's a checkbox
		Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" " & arrCB(checkPointsCount-1) & " >  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
												'arrCB(checkPointsCount-1)
	ELSE 
		Response.Write "<td contenteditable             onclick=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")""                     ></td> " 'when row number is 2, read from cell NUMBER 5 through 8, when cell number is 6 , it's a checkbox
		Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" unchecked >  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
												'arrCB(checkPointsCount-1)
	END IF

End If

IF checkPointsCount>1 Then 'starting from number 3, read from cell number 2 through 5 and when cell number is 3 it's a checkbox
	Response.Write "<tr> <td   >" &  " <a href=""#"" onmouseenter=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")"" >"  &  objRS.Fields("checkpoint") &  "</a> "  &  "</td>  <td   > " & objRS.Fields("fulfill_standard") & "</td> "
	IF loadInputHistory = "Y" THEN 
		Response.Write "<td  contenteditable    onclick=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")""         >" &  arr(checkPointsCount-1) & "</td> " 
		Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" " & arrCB(checkPointsCount-1) & " >  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
	ELSE 
		Response.Write "<td  contenteditable      onclick=""checkPointClicked(" & objRS.Fields("checkpoint_id") & ")""           ></td> " 
		Response.Write "<td >  <input type=""checkbox""  id=""" & checkPointsCount &  "000"" name=""vehicle"" value=""Car"" unchecked >  </td>   <td style=""display:none;"">" & objRS.Fields("sub_cat_id") & "</td> <td style=""display:none;"">" & objRS.Fields("checkpoint_id") & "</td> </tr>"
    END IF	
End If
checkPointsCount=checkPointsCount+1
objRS.MoveNext
Wend
Response.Write "</table>"
Response.Write "<br>"
'Response.Write startCheck 
Response.Write buttonBack
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
	//var checkPointsIdlistValue = getInlist(audit_procedures);
    var issueListValue = document.getElementById('issueList').innerHTML;
	var parameters = {
	  checkPointIdList: getInlist(audit_procedures),
	  procedureList: procedureListBuilder(audit_procedures),
	  issueList: issueListValue
	}


	for (var name in parameters) {
	  if (parameters.hasOwnProperty(name)) {
		console.log('this is fog (' + name + ') for sure. Value: ' + parameters[name]);
	  }
	  else {
		console.log(name); // toString or something else
	  }
	}
	post("issueTrackerBuilder.asp",parameters)
	//navigateToIssueTracker(checkPointsInlist);
}

function post(path, params, method) {
    method = method || "post"; // Set method to post by default if not specified.

    // The rest of this code assumes you are not using a library.
    // It can be made less wordy if you use one.
    var form = document.createElement("form");
    form.setAttribute("method", method);
    form.setAttribute("action", path);
	
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

function navigates(navigateTo){
 window.location.href = navigateTo
}

function destroySessionVar(varName){
var http = new XMLHttpRequest();
var url = "destroySessionVar.asp";
varName="varName=" + varName
http.open("POST", url, true);
http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
http.onreadystatechange = function(){//Call a function when the state changes.
    if(http.readyState == 4 && http.status == 200){
       alert(http.responseText);  //debug
     }
   }
http.send(varName);
}

function homeClicked(){
navigates("home.asp");
}

function navigates(navigateTo){
 window.location.href = navigateTo;
}

function buttonBack(){
	if (document.referrer.includes("selfEvaNavigation")){
		navigates(document.referrer);
	} else {
	   	var singleCatSheetPgData 	= document.getElementById("singleCatSheetPgData");
		var chkType  				= adminPageData.getAttribute('chkType');
		var account  				= adminPageData.getAttribute('account');
	    var params					= "reqType=" + "selfEva" + "&" +
									= "chkType=" + chkType + "&" +
									= "account=" + account;					
	   navigates("selfEvaNavigation.asp?" + params);
	}
}

</script>
 
</body>
</html>