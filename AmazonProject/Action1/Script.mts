SystemUtil.CloseProcessByName"chrome.exe" @@ hightlight id_;_12322162_;_script infofile_;_ZIP::ssf13.xml_;_
SystemUtil.Run"Chrome.exe","www.amazon.in"
 @@ script infofile_;_ZIP::ssf122.xml_;_
On Error Resume Next
FilePath="C:\Users\user262\Documents\AmazonProject\Test Data.xlsx"
ExcelSheet= "Test Data"
SheetName="Login Data"
Datatable.AddSheet ExcelSheet
DataTable.ImportSheet FilePath,SheetName,ExcelSheet

rowCount=Datatable.GetSheet(ExcelSheet).GetRowCount

For i=1 To rowCount
DataTable.SetCurrentRow(i)
If DataTable("Execution_Flag",ExcelSheet)="Y" Then
ExecuteTest (DataTable.Value("TC_ID",ExcelSheet))
DataTable.Value("Result",ExcelSheet)=Environment.Value("Result")
End if
Next
DataTable.ExportSheet FilePath,ExcelSheet,SheetName


 
 @@ script infofile_;_ZIP::ssf128.xml_;_
 @@ script infofile_;_ZIP::ssf120.xml_;_
 @@ script infofile_;_ZIP::ssf132.xml_;_
 @@ script infofile_;_ZIP::ssf139.xml_;_
