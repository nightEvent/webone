<%
response.expires=-1
response.ContentType = "text/plain"
Response.CharSet = "UTF-8"
Session.CodePage = 65001
Session.LCID     = 2052 

Dim setID,chkType,subCat,objConn,sqlInsertCheckPoint,sConnection,objRS
setID=Request.Form("setID")
chkType=Request.Form("chkType")
subCat=Request.Form("subCat")
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(sConnection)
'retrieve from database, wrap it into a view
'http://stackoverflow.com/questions/33369581/how-to-calculate-ratio-in-mysql
'database rows to d3
'http://www.d3noob.org/2013/02/using-mysql-database-as-source-of-data.html

sqlReport = " select gs.set_name as set_name, gs.total/s.total as frequency" 	& _
			"	FROM"										   	& _												
			"	(   "										   	& _
			"	 select count(*) as total, set_name "			& _		 
			"	   from SumByTime"								& _
			"	  where chk_Type = 'Z'"						   	& _
			"		AND creation_time > TIMESTAMPADD(DAY,-7, Now() ) " & _
			"	  group by set_name "								& _
			"	) gs "											& _
			"	, 	 "											& _
			"	(	 "											& _
			"	 select count(*) as total"						& _
			"	   from SumByTime "								& _					
			"	  where chk_Type = 'Z'"							& _
			"	   AND creation_time > TIMESTAMPADD(DAY, -7, Now() )" & _
			"	) s "


Set objRS = objConn.Execute(sqlReport)
'sConnection.CommitTrans
Dim ind,tab,newLine
ind=1
tab= Chr(9)
newLine=Chr(10)
While Not objRS.EOF
if ind = 1 then
response.write "set_name" & tab & "frequency" & newLine
response.write objRS.Fields("set_name")  & tab & objRS.Fields("frequency") & newLine
else
response.write objRS.Fields("set_name")  & tab & objRS.Fields("frequency") & newLine
end if
ind=ind+1
objRS.MoveNext
Wend


'While Not objRS.EOF
'if ind = 1 then
'response.write "[{set_name:""" & objRS.Fields("set_name") & """,frequency:""" & objRS.Fields("frequency") & """}"
'else
'response.write ",{set_name:""" & objRS.Fields("set_name") & """,frequency:""" & objRS.Fields("frequency") & """}"
'end if
'objRS.MoveNext
'Wend
'response.write "]" 

Set objRS = Nothing
objConn.Close
Set objConn = Nothing
'objconn.CommitTrans
%>
