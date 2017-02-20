<!DOCTYPE html>
<html>
<%@ language="VBScript" codepage="65001" %>
<%
Response.CharSet = "utf-8"
session.codepage=65001
%>

<head>

<title>hello..</title>

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



</head>

<body>

<%
'Start the session and store information
Function Lpad (sValue, sPadchar, iLength)
  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
End Function
Dim currentTime
Session("TimeVisited") = Now()
currentTime=Now()
yyy=Year(Date())
yyy= RIGHT(yyy, (LEN(yyy)-2))
mon=Month(currentTime)
mon=Lpad(mon,"0",2)
ddd=Day(currentTime)
ddd=Lpad(ddd,"0",2)
hhh=Hour(currentTime)
hhh=Lpad(hhh,"0",2)
mmm=Minute(currentTime)
mmm=Lpad(mmm,"0",2)
sss=Second(currentTime)
sss=Lpad(sss,"0",2)
Response.Write("You Date() is: " & Date() & ". </br> ")
Response.Write("You Now() is: " & Now() & ". </br> ")
Response.Write("You Now() is: " & yyy & ddd & hhh & mmm & sss & "  . </br>")
timeString=yyy & ddd & hhh & mmm & sss
Response.Write("You yearToSecond is: " & timeString & " </br>")
Response.Write("You yearToSecond + 1  is: " & ( timeString + 1 )  & " </br>")





%>

<%
dim i
For Each i in Session.Contents
  Response.Write(i & "<br>")
Next
dim ii
dim jj
jj=Session.Contents.Count
Response.Write("Session variables: " & jj)
For ii=1 to jj
  Response.Write(Session.Contents(ii) & "<br>")
Next
Session.Timeout=1
'Abandon method to end a session immediatel
'Session.Abandon
%>

<button onclick="search()">click to expand self evaluate sheet</button>

<script>

function search(){
   document.write ( "<div id=" + "sheet"+ "> .</div>")
   document.getElementById("sheet").style.border = "thin solid black"
 
   searchStr="1"    //sub_cat_id 

   if (searchStr.length == 0) { 
        document.getElementById("sheet").innerHTML = "  ";
        return;
    } else {
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (this.readyState == 4 && this.status == 200) {
                document.getElementById("sheet").innerHTML = this.responseText;
				
				//lightOn();
            }
        };
        xmlhttp.open("GET", "selfEvaSheetBuilder.asp?q="+searchStr, true); 
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

function saveInDB(audit_procedures){
alert('saving in Database');
}

function navigateToIssueTracker(inlist){
var navigateTo="http://localhost:88/issueTrackerBuilder.asp?checkPointIdList=" + inlist;
//alert('navigate to ' + navigateTo);
 window.location=navigateTo;
}; 

function startChecking(){
var table = document.getElementById('selfEvaSheet');
var checkPointsCount = document.getElementById('checkPointsCount').value;
var audit_procedures = createArray(checkPointsCount, 4 ); 
//user hide button 排查开始, 
//show the hiden cells 排查经过 and 发现问题, 
//then we have data input by user, 
//so we call getCellValues(audit_procedures) to collect the data, 
//and we create a form with a 提交排查经过 submit button in it to let user decide if they want to post the data into the database
//buttonSubmit="<button onclick=""startChecking()""> 提交 </button>"
getCellValues(audit_procedures,table);
//alert('print');
//printArrays(audit_procedures);
//saveInDB(audit_procedures);
var checkPointsInlist
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
checkPointsInlist = getInlist(audit_procedures);
//alert(checkPointsInlist);
navigateToIssueTracker(checkPointsInlist);
}

</script>


</body>
</html>