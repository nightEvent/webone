<%
Session.Timeout=1440
response.expires=-1
Dim sConnection, objConn , objRS ,headerRow, queryStr, hostname,catCount,delimiter,sourceRequestType,setID,subCatID,checkPointType,subCatReturned

sourceRequestType=Request.Form("RequestType")
'Response.write Request.Form("RequestType")
if sourceRequestType = "AR" then
	subCatID=Request.Form("subCatID")
	checkPointType=Request.Form("checkPointType")
	queryStr=" select audit_rule from webone.audit_method where sub_cat_id = " & _
	         subCatID & "  and checkpoint_type = " & "'"  & checkPointType & "' ; "
'response.write "AR:queryStr:" & queryStr
	sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
	Set objConn = Server.CreateObject("ADODB.Connection") 
	objConn.Open(sConnection) 
	Set objRS = objConn.Execute(queryStr)
	catCount = 1 
	delimiter = Chr(31)

	While Not objRS.EOF
			if catCount = 1 then
				response.write objRS.Fields("sub_cat_id") & delimiter & objRS.Fields("sub_cat_name")
			else 
				response.write delimiter & objRS.Fields("sub_cat_id") & delimiter & objRS.Fields("sub_cat_name")
			end if
			catCount = catCount + 1
		objRS.MoveNext
	Wend
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing
elseif  sourceRequestType = "SC" then
	setID=Request.Form("setID")
	checkPointType=Request.Form("checkPointType")
	queryStr=" select sub_cat_id, sub_cat_name from webone.sub_category where set_id = " & setID & _
             " AND chk_Type = '" & checkPointType &  "' ; "
	sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
	Set objConn = Server.CreateObject("ADODB.Connection") 
	objConn.Open(sConnection) 
	Set objRS = objConn.Execute(queryStr)
	catCount = 1 
	delimiter = Chr(31)
	subCatReturned="NoSubCatReturn"
	While Not objRS.EOF
	    subCatReturned="HaveSubCatReturn"
		if catCount = 1 then
			response.write objRS.Fields("sub_cat_id") & delimiter & objRS.Fields("sub_cat_name")
		else 
			response.write delimiter & objRS.Fields("sub_cat_id") & delimiter & objRS.Fields("sub_cat_name")
		end if
		catCount = catCount + 1
		objRS.MoveNext
	Wend
	if subCatReturned="NoSubCatReturn" then
	   response.write "NoSubCatReturn"
	end if
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing
end if


%>
