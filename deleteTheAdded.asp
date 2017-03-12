<%
Response.CharSet = "utf-8"
response.expires=-1
Session.Timeout=1440

Dim sConnection, objConn , objRS,deleteType , deleteStr ,checkPointsList,subCatList
deleteType		= Request.Form("deleteType")
if deleteType="C" then
    checkPointsList = Request.Form("checkPointIdList")
	deleteStr="DELETE FROM webone.checkpoints WHERE checkpoint_id in (" & checkPointsList & ");"
elseif deleteType="S" then
	subCatList		= Request.Form("subCatList")
	deleteStr="DELETE FROM webone.sub_category WHERE sub_cat_id in (" & subCatList & ");"
elseif deleteType="L" then											'NOT YET DONE reserved for linked rule  
	subCatList		= Request.Form("subCatList")                    'NOT YET DONE reserved for linked rule 
	checkPointsList = Request.Form("checkPointIdList")				'NOT YET DONE reserved for linked rule 
	deleteStr="DELETE FROM webone.checkpoints WHERE checkpoint_id in (" & checkPointsList & ");"  'NOT YET DONE
end if
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(deleteStr)
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
Response.Write "选择的项目已删除！"
%>
