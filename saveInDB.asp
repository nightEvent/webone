<%
Response.CharSet = "utf-8"
response.expires=-1
Session.Timeout=1440
Function Lpad (sValue, sPadchar, iLength)
  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
End Function

Function KeyGenerator(baseTime,interval,salt)
'Response.Write("baseTime before intervalled: " & baseTime & ". </br> ")
baseTime=DateAdd("s",interval,baseTime)
'Response.Write("baseTime after intervalled: " & baseTime & ". </br> ")
yyy=Year(Date())
yyy= RIGHT(yyy, (LEN(yyy)-2))
mon=Month(baseTime)
mon=Lpad(mon,"0",2)
ddd=Day(baseTime)
ddd=Lpad(ddd,"0",2)
hhh=Hour(baseTime)
hhh=Lpad(hhh,"0",2)
mmm=Minute(baseTime)
mmm=Lpad(mmm,"0",2)
sss=Second(baseTime)
sss=Lpad(sss,"0",2)
timeString=yyy & mon & ddd & hhh & mmm & sss
timeString=(timeString*1000 + salt)
KeyGenerator=timeString
End Function

'generate insert procedures statement starts
Dim procedureList,issueList,recordsArr(),keyIdArr(),ind,procedureExpired,currentTime,inList,currentTimeDelt,keyCurrent,setId
procedureList=Request.Form("procedureList")
issueList=Request.Form("issueList")
setId=Request.Form("setId")
If procedureList = "" or issueList = "" Then
   procedureExpired="N"
else
   procedureExpired="Y"
End If


delimiter = Chr(31)
ind=0
recordInd=0

'get current year and time
currentTime=Now()
arrays=Split(procedureList,delimiter)
recordCnt = uBound(arrays)/3
ReDim Preserve recordsArr(recordCnt-1)
ReDim Preserve keyIdArr(recordCnt-1,1)

'######
inList = "feelSoDownToday"
'Response.Write(" HLN " & currentTime)
'loop through audit procedure
if procedureExpired = "Y" Then
	for each x in arrays
		If (ind mod 3) = 2 Then
			if inList = "feelSoDownToday" Then
				inList =  x 
			else
				inList = inList & ","  & x 
			end if
			currentTimeDelt=currentTime
			keyCurrent = KeyGenerator(currentTimeDelt,(recordInd - recordCnt),setId)
		    recordsArr(recordInd) = keyCurrent & "," & inList
			keyIdArr(recordInd,0) = x
			keyIdArr(recordInd,1) = keyCurrent
			recordInd = recordInd + 1
			inList = "feelSoDownToday"
		ELSE 
		   if inList = "feelSoDownToday" Then
				inList = """" & x & """"
			else
				inList = inList & "," & """" & x & """"
			end if
		End If
		ind= ind + 1
	next
else
Response.Write "抱歉，数据已经丢失请联系管理员！"
Response.End
End If


Dim procedureInsert
procedureInsert="thisTimeFeelGreat"
for each x in recordsArr
if procedureInsert="thisTimeFeelGreat" then 
procedureInsert=" insert into webone.audit_procedure (audit_id, procedures,issue_found_flag,checkpoint_id) values " 
procedureInsert= procedureInsert & "(" & x & ")"
else 
procedureInsert= procedureInsert & ",(" & x & ")"
end if
next
procedureInsert= procedureInsert + "  ;"
'-------- generate insert inssues statements starts ----------------
Dim issuesArr(),issueInd,hasIssue,issuesCnt

delimiter = Chr(31)
issueInd=0
recordInd=0
hasIssue="N"
arraysDelt=Split(issueList,delimiter)
issuesCnt = uBound(arraysDelt)/4
ReDim Preserve issuesArr(issuesCnt-1)
inList = "feelSoFullNow"
'loop through audit procedure
if issueList <> "" Then
	hasIssue="Y"
	for each x in arraysDelt
		If (issueInd mod 4) = 3 Then
		  For i = 0 To recordCnt
		  	 if keyIdArr(i,0) = x then
				inList = inList & "," & keyIdArr(i,1)
			 end if
		  Next
		  issuesArr(recordInd)=inList
		  recordInd = recordInd + 1
		  inList = "feelSoFullNow"
		ELSE 
		   if inList = "feelSoFullNow" Then
				inList = """" & x & """"
			else
				inList = inList & "," & """" & x & """"
			end if
		End If
		issueInd= issueInd + 1
	next
End If
Dim issuesInsert
issuesInsert="thisTimeFeelGreat"
for each x in issuesArr
if issuesInsert="thisTimeFeelGreat" then 
issuesInsert=" insert into webone.issue_tracking (issue, corrections,result,audit_id) values " 
issuesInsert= issuesInsert & "(" & x & ")"
else 
issuesInsert= issuesInsert & ",(" & x & ")"
end if
next

issuesInsert= issuesInsert + "  ;"
'response.write "issuesInsert is " & issuesInsert
'generate insert inssues statements ends
Dim sConnection, objConn , objRS
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(sConnection)
Set objRS = objConn.Execute(procedureInsert)
Set objRS = objConn.Execute(issuesInsert)
'
'sConnection.CommitTrans
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
Response.Write "数据已提交，点击确定回到系列页面。"
'objconn.CommitTrans
%>