<%
Response.CharSet = "utf-8"
response.expires=-1
Session.Timeout=1440

Dim checkPointType,subCatID,auditRuleEitherNewOrUpdated,checkPointContent,checkPointFulfillStandard
Dim checkPointInsert,sqlAuditRulePresent,updateAuditRule,sConnection, objConn , objRS
Dim sqlUpdateAuditRule,sqlInsertAuditRule,sqlInsertCheckPoint,auditRuleChanged
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(sConnection)

checkPointType=Request.Form("checkPointType")
subCatID=Request.Form("subCatID")
auditRuleEitherNewOrUpdated=Request.Form("auditRuleEitherNewOrUpdated")
checkPointContent=Request.Form("checkPointContent")
checkPointFulfillStandard=Request.Form("checkPointFulfillStandard")
auditRuleChanged=Request.Form("auditRuleChanged")

if checkPointType = "L" OR checkPointType = "Z" then 'only when audit rule exist thus need to insert or update, otherwise insert checkpoint and fulfill standard
	'check existence of audit rule
	if auditRuleChanged = "Y" then
		sqlAuditRulePresent = "SELECT 'Y' as auditRulePresent FROM webone.audit_method WHERE sub_cat_id = " & subCatID & " and checkpoint_type = '" & checkPointType & "' ;"
		Set objRS = objConn.Execute(sqlAuditRulePresent)
		updateAuditRule="N"
		While Not objRS.EOF
			updateAuditRule = objRS.Fields("auditRulePresent")  
			objRS.MoveNext
		Wend
		'update or insert for audit rule
		if updateAuditRule = "Y" then
			sqlUpdateAuditRule = "UPDATE webone.audit_method SET audit_rule = " & " '" & auditRuleEitherNewOrUpdated & "' " & _
			" WHERE sub_cat_id = " & subCatID & " and checkpoint_type = " & "'" & checkPointType & "' ;"
			Set objRS = objConn.Execute(sqlUpdateAuditRule)
		else
			sqlInsertAuditRule = "INSERT INTO webone.audit_method (audit_rule,checkpoint_type,sub_cat_id) VALUES ( '" & auditRuleEitherNewOrUpdated & "'" & _
			",'" & checkPointType & "'" & _
			"," & subCatID & ");"
			Set objRS = objConn.Execute(sqlInsertAuditRule)
		end if
	end if
end if
'insert for checkpoint
sqlInsertCheckPoint = "INSERT INTO webone.checkpoints (sub_cat_id,content,fulfill_standard,checkpoint_type) VALUES ( " & subCatID & _
", '" & checkPointContent & "'" & _
", '" & checkPointFulfillStandard & "'" & _
", '" & checkPointType & "'"  & _
");"
Set objRS = objConn.Execute(sqlInsertCheckPoint)
'sConnection.CommitTrans
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
response.write "Y"
'objconn.CommitTrans
%>