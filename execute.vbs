Dim Arg, FilePath
Set Arg = WScript.Arguments

'Usage Parameters
if WScript.Arguments.Count <> 1 then
    WScript.Echo "Missing parameter"
    WScript.Echo "Usage: execute.vbs " & Chr(34) & "Path to Excel document.xlsx" & Chr(34)
    WScript.Quit 1
end if

' Set the File path of the excel document
FilePath = Arg(0)

' Excel Macro logic
Const xlCellTypeBlanks = 4

Dim xlApp
Dim xlwb
Dim countSheet

Set xlApp = CreateObject("Excel.Application")
Set xlwb = xlApp.workbooks.Open(FilePath)
'Find the number of sheets
countSheet = xlwb.Sheets.Count

'Run a loop to the last sheet 
On Error Resume Next
For i = 1 to countSheet
xlwb.Sheets(i).Columns("a:a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
On Error GoTo 0

Next
'Save and Close the Excel document and the Application
xlwb.Save
xlwb.Close
xlApp.Quit

WScript.Echo "Empty Rows Deleted"

'Clear the objects at the end of your script.
set Arg = Nothing
