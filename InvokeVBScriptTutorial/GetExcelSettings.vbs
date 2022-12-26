'Get Excel Settings
On Error Resume Next
GetExcelSeparator
If Err.Number <> 0 Then
	WScript.stdout.WriteLine(Err.Description)
	CloseExcel
End If



Function GetExcelSeparator
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	objExcel.DisplayAlerts = False
	
	objExcel.DecimalSeparator = "|"

	 vResult = objExcel.DecimalSeparator
	 WScript.stdout.WriteLine(vResult)
	 Set objExcel = Nothing
end Function
 
Function CloseExcel
	Set objExcel = GetObject(, "Excel.Application")  'attach to running Excel instance
	Set wb = Nothing
	For Each obj In obj.Workbooks
		obj.Quit
	Next
	Set objExcel = Nothing
end Function

