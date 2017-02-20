<%
Response.ContentType = "application/msword"
Response.AddHeader "Content-Disposition", "attachment;filename=甘肃分公司内控风险点自查工作报告.doc" 
Response.write("<html " & _ 
"xmlns:o='urn:schemas-microsoft-com:office:office' " & _
"xmlns:w='urn:schemas-microsoft-com:office:word'" & _ 
"xmlns='http://www.w3.org/TR/REC-html40'>" & _
"<head> <meta charset=""UTF-8"">" ) 

'The setting specifies document's view after it is downloaded as Print
'instead of the default Web Layout
Response.write("<!--[if gte mso 9]>" & _
"<xml>" & _ 
"<w:WordDocument>" & _
"<w:View>Print</w:View>" & _
"<w:Zoom>90</w:Zoom>" & _ 
"<w:DoNotOptimizeForBrowser/>" & _
"</w:WordDocument>" & _
"</xml>" & _ 
"<![endif]-->")

' Write mso style definitions class in style tag, style class can be divided by
' sections, in this example, i used Section1
' Then call your style class (Section1) in div tag

Response.write("<style>" & _
"<!-- /* Style Definitions */" & _
"@page Section1" & _
" {size:8.5in 11.0in; " & _
" margin:1.0in 1.25in 1.0in 1.25in ; " & _
" mso-header-margin:.5in; " & _
" mso-footer-margin:.5in; mso-paper-source:0;}" & _
" div.Section1" & _
" {page:Section1;}" & _
"-->" & _
"p { " & _
"  font-family: ""FangSong_GB2312"", Times, serif; " & _
"  font-size: 16pt; " 	& _
"}"  					& _
"</style></head>") 
Response.write("<body lang=UTF-8 style='tab-interval:.5in'>" 	& _
"<div class=Section1>" 											& _
"<h1><center>太平人寿甘肃分公司关于2016年度内控</center></h1>"  & _ 
"<h2><center>风险点自查工作的报告</center></h2>" 				& _ 
"<p >" 															& _
"分公司：<br>" 													& _
" &nbsp 根据总经理室领导工作指示，2016年8月，由分公司 "   		& _
"由分公司风险管理及合规部（监察室）牵头，各部门、各条线参与，"  & _
"内控风险点自查工作在全省各机构落实开展。9月初，该项"   		& _
"工作进入总结阶段。现将本次分公司内控风险点自查工作开"   		& _
"展情况汇报如下：<br> "   										& _
"一、自查工作开展过程 <br> "  									& _
  "（一）准备阶段（7月15日-7月31日）<br> "   					& _
"根据《总公司合规制度汇编》（）、总公司下发至分公司"  		 	& _
"总公司下发至分公司各部门的最新风险点目录、及全省各单位实际架构经营情况，"   & _
"分公司风险管理及合规部（监察室）多次召开讨论会议，按"   & _
"部门条线整理完成《太平人寿甘肃分公司内控风险自查表》，" & _
"以更加科学、高效开展本次内控风险点自查工作。<br> "   	 & _
  "（二）自查工作执行阶段（8月9日-8月25日） <br> "   	 & _
"2016年8月，分公司下发工作通知书，2016年度内控"   	     & _
"风险点自查工作正式开始执行。各部门、各机构根据《太平"   & _
"人寿甘肃分公司内控风险自查表》中的具体内容，各部门、"   & _
"各条线对应相应风险点，认真开展本次工作。<br> "   		 & _
" &nbsp 太平人寿甘肃分公司内控风险自查表》将自查工作项"  & _
"目具体细化至个险、教培、银保、保费、财务、运营、人力"   & _
"资源、办公室、企划及合规各个部门条线，一共涉及210项"    & _
"风险点。表内内容包含具体风险点描述、自查方法及可参考"   & _
"制度。各部门条线需填写反馈具体风险点对应的排查结果。"   & _
"排查结果包含发现问题、整改措施和预计整改完成时间。该"   & _
"表格既是本次内控风险点自查工作的核心工具，也是分公司"   & _
"各部门、各机构条线开展工作的重要方法指引。<br>"   		 & _
" &nbsp （三）工作反馈阶段（8月26日-8月31日 <br> "       & _
" &nbsp 分公司各部门将排查结果汇总至《太平人寿甘肃分公司"   & _
"内控风险自查表》内，各机构合规联系人汇总各条线排查结 "   & _
"果。各部门、各机构按时将汇总结果反馈至分公司风险管理"    & _
"及合规部（监察室）。 <br>"   							  & _
" &nbsp 二、自查工作发现问题及整改情况 <br>" )
'################# middle to feed selfevaluations
Dim sConnection, objConn , objRS ,queryStr
queryStr="SELECT set_name,sub_cat_name,issue,corrections,result FROM webone.reporting where 1= 1 order by set_id asc  ;"
sConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=localhost; DATABASE=webone; UID=weboneuser;PASSWORD=weboneuser;PTION=3" 
Set objConn = Server.CreateObject("ADODB.Connection") 
objConn.Open(sConnection) 
Set objRS = objConn.Execute(queryStr)

Function chineseOneDigit(arabicNumeral)
select case arabicNumeral
  case 1 
     chineseOneDigit="一"
  case 2
     chineseOneDigit="二"
  case 3
     chineseOneDigit="三"
  case 4 
     chineseOneDigit="四"
  case 5
     chineseOneDigit="五"
  case 6
     chineseOneDigit="六"
  case 7 
     chineseOneDigit="七"
  case 8
     chineseOneDigit="八"
  case 9
     chineseOneDigit="九"
  case else 
     chineseOneDigit="零"
end select
End Function

Function chineseNum(arabicNumeral) 
if arabicNumeral < 9 then
 chineseNum=chineseOneDigit(arabicNumeral)
elseif (arabicNumeral > 9 and arabicNumeral < 100) then
chineseNum=chineseOneDigit(Int(arabicNumeral/10)) & "十" & chineseOneDigit(arabicNumeral mod 10)
else
chineseNum="数字越界"
end if
End Function

previousSet="notSet"
setIndex=1
previousSubCat="notSet"
subCatIndex=1
issueIndex=1
While Not objRS.EOF
if objRS.Fields("set_name") = previousSet then

	if objRS.Fields("sub_cat_name") = previousSubCat then
	   Response.Write "(" & issueIndex & ")" & "问题描述：" & objRS.Fields("issue") & " <br> "
	   Response.Write " &nbsp  &nbsp 整改措施：" & objRS.Fields("corrections") & " <br> "
	   Response.Write " &nbsp  &nbsp 预计完成时间：" & objRS.Fields("result") & " <br> "
	   issueIndex=issueIndex+1
	else
	   issueIndex=1
	   Response.Write " &nbsp  &nbsp " & subCatIndex & "." & objRS.Fields("sub_cat_name") & " <br> "
	   Response.Write "(" & issueIndex & ")" & "问题描述：" & objRS.Fields("issue") & " <br> " 
	   Response.Write " &nbsp  &nbsp 整改措施：" & objRS.Fields("corrections") & " <br> "
	   Response.Write " &nbsp  &nbsp 预计完成时间：" & objRS.Fields("result") & " <br> "
	   subSetIndex=subCatIndex + 1
	   issueIndex=issueIndex + 1
	   previousSubCat=objRS.Fields("sub_cat_name")
	end if

else
    subCatIndex=1
	issueIndex=1
	Response.Write "（" & chineseNum(setIndex) & "）" &  objRS.Fields("set_name") & "条线 <br> "
	Response.Write subCatIndex & "." & objRS.Fields("sub_cat_name") & " <br>"
	Response.Write "(" & issueIndex & ")" & "问题描述：" & objRS.Fields("issue") & " <br> " 
	Response.Write " &nbsp  &nbsp 整改措施：" & objRS.Fields("corrections") & " <br> "
	Response.Write " &nbsp  &nbsp 预计完成时间：" & objRS.Fields("result") & " <br> "
	previousSet=objRS.Fields("set_name")
	previousSubCat=objRS.Fields("sub_cat_name")
    subCatIndex=subCatIndex + 1
	issueIndex= issueIndex + 1
	setIndex=setIndex + 1
end if
objRS.MoveNext
Wend
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing


'#################ending
Response.write ( "  &nbsp   &nbsp  三、下一步工作计划 <br> " )
Response.write ( "  &nbsp   &nbsp  各部门、各机构将按照计划时间，认真落实整改措施，" )
Response.write ( "对排查出的问题按时完成整改。同时，各机构将加强与分公" )
Response.write ( "沟通，协商部分疑难及搁置问题的整改方法。各部门、各机" )
Response.write ( "构将同心协力，进一步巩固合规意识、为甘肃分公司健康稳" )
Response.write ( "健发展共同助力。<br> " )
Response.write ( "  &nbsp   &nbsp  特此报告<br> " )
Response.write ( " <br> " )
Response.write ( " <br> " )
Response.write ( "&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp 风险管理及合规部（监察室）<br> " )
baseTime = Now()
yyy=Year(Date())
mon=Month(baseTime)
ddd=Day(baseTime)
'Response.Write("baseTime after intervalled: " & baseTime & ". </br> ")
Response.write ( "&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp  " & yyy & "年" &  mon & "月" & ddd & "日")

Response.write ( " </p>" & _
"</div></body>")
Response.write ("</html>") 

%> 
