vNumber = WScript.Arguments.Item(0)
'vNumber = 4.6449200614
vNumber = Round(vNumber,2)
vPos = InStr(vNumber,".")
if vPos = 0 Then
	vNumber = Cstr(vNumber) + "." + "00"
Else
	vLen = len(vNumber)
	vLenAfterDec = Mid(vNumber,(vPos+1))
	if len(vLenAfterDec) = 1 Then
		vNumber = Cstr(vNumber) + "0"
	End if
End if
WScript.StdOut.WriteLine vNumber