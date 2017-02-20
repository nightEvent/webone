<%@ language="VBScript" codepage="65001" %>
<%
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWB = objExcel.Workbooks.Add
Set objSheet = objWB.Sheets(1)

While objWB.Worksheets.Count < 13
objWB.Worksheets.Add ,objWB.Worksheets(objWB.Worksheets.Count)
Wend

objWB.Worksheets(1).Name = "Cover"
objWB.Worksheets(2).Name = "Headquarters"
objWB.Worksheets(3).Name = "Miami"
objWB.Worksheets(4).Name = "Chicago"
objWB.Worksheets(5).Name = "Dallas"
objWB.Worksheets(6).Name = "6"
objWB.Worksheets(7).Name = "7"
objWB.Worksheets(8).Name = "8"
objWB.Worksheets(9).Name = "9"
objWB.Worksheets(10).Name = "10"
objWB.Worksheets(11).Name = "11"
objWB.Worksheets(12).Name = "12"
objWB.Worksheets(13).Name = "13"

Set objSheet = objWB.Worksheets(2)
 
objSheet.Cells(1, 1) = "Computer Name"
objSheet.Cells(1, 2) = "Make"
objSheet.Cells(1, 3) = "Model"
objSheet.Cells(1, 4) = "Serial Number"
objSheet.cells(1, 5) = "Operating System"
objSheet.cells(1, 6) = "Service Pack"
objSheet.cells(1, 7) = "Image Date"
objSheet.Rows("1:1").Font.Bold = True

objSheet.Activate

%>