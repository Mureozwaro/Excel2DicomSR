Dim xlApp, xlBook
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True

filePath = "\\childrens\files\WebForms\CTA\CTA_Cardiac.xlsm"

If WScript.Arguments.Count <> 5 Then
    WScript.Echo "Usage: Launch.VBS <arg1> <arg2> <arg3> <arg4> <arg5>"
    WScript.Quit 1
Else 
   arg1 = WScript.Arguments(0)
   arg2 = WScript.Arguments(1)
   arg3 = WScript.Arguments(2)
   arg4 = WScript.Arguments(3)
   arg5 = WScript.Arguments(4)
End If

Set objWorkbook = xlApp.Workbooks.Open(filePath)

Set objSheet = objWorkbook.Sheets("Z Value Calculator")

' Patient Name (1, 2), MRN (3, 2), Accession (5,2), Study Date (7,2), User ID (1, 9)

objSheet.Cells(1, 2).Value = arg1
objSheet.Cells(3, 2).Value = arg2
objSheet.Cells(5, 2).Value = arg3
objSheet.Cells(7, 2).Value = arg4
objSheet.Cells(1, 9).Value = arg5

Set xlBook = Nothing
Set xlApp = Nothing