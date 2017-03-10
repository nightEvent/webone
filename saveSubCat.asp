<%
Response.CharSet = "utf-8"
response.expires=-1
Session.Timeout=1440
Dim setID,chkType,subCat,objConn,sqlInsertCheckPoint,sConnection,objRS
setID=Request.Form("setID")
chkType=Request.Form("chkType")
subCat=Request.Form("subCat")
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(sConnection)

'insert for checkpoint
sqlInsertCheckPoint = "INSERT INTO webone.sub_category (set_id,chk_Type,sub_cat_name) VALUES ( " & setID  & _
", '" & chkType & "'" & _
", '" & subCat & "'" & _
");"
Set objRS = objConn.Execute(sqlInsertCheckPoint)
'sConnection.CommitTrans
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
response.write "Y"
'objconn.CommitTrans
%>