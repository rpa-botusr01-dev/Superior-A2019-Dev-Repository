Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
on error resume next
Set objWorkBookSrcFile = objExcel.Workbooks.Open(WScript.Arguments.Item(0))
Set objWorkBookSrc = objWorkBookSrcFile.Worksheets("Data")
FileLastRow = objWorkBookSrc.UsedRange.Rows.Count

objWorkBookSrc.Cells(1, 24).Value = "Open Amount"
objWorkBookSrc.Cells(1, 25).Value = "Amount To Apply"
For i = 2 To FileLastRow
	objWorkBookSrc.Cells(i, 24).Value = objWorkBookSrc.Cells(i, 5).Value - objWorkBookSrc.Cells(i, 23).Value
Next

objWorkBookSrcFile.Save
objWorkBookSrcFile.Close
objExcel.Quit

