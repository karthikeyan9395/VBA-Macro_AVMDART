
'**********************************************************************************************************************'
' Two things hardcoded here.
' 1. In Sheet1 the values needs to be started from (2,11) i.e) K cell
' 2. In sheet2 the values needs to be started from (2,3) i.e) C cell
' Sheet names are hard coded as Sheet1 and Sheet2
' If there is any empty row in the top remove it.
'**********************************************************************************************************************'

Dim dbsheet1 As Worksheet
Dim dbsheet2 As Worksheet

Sub Main()

Dim r1, c1, r2, c2 As Integer

Set dbsheet1 = ActiveWorkbook.Sheets("Sheet1")
Set dbsheet2 = ActiveWorkbook.Sheets("Sheet2")

dbsheet1.Activate
r1 = dbsheet1.Cells(Rows.Count, 1).End(xlUp).Row
c1 = dbsheet1.Cells(1, Columns.Count).End(xlToLeft).Column

dbsheet2.Activate
r2 = dbsheet2.Cells(Rows.Count, 1).End(xlUp).Row
c2 = dbsheet2.Cells(1, Columns.Count).End(xlToLeft).Column

Application.CutCopyMode = False
dbsheet2.Range(Cells(1, 3), Cells(1, c2)).Copy
dbsheet1.Activate
dbsheet1.Cells(1, c1 + 1).Select
dbsheet1.Cells(1, c1 + 1).PasteSpecial
'dbsheet1.Range(Cells(1, c1 + 1), Cells(1, c2)).PasteSpecial

c1 = dbsheet1.Cells(1, Columns.Count).End(xlToLeft).Column

Table1 = dbsheet1.Range(Cells(2, 5), Cells(r1, 5)) 'Name Column from Original
dbsheet2.Activate
Table2 = dbsheet2.Range(Cells(2, 1), Cells(r2, c2)) ' Range of Table 2 for Vlookup

dbsheet1.Activate
Row = 2
a = 3

For Each x In Table1

    For i = 11 To c1
        dbsheet1.Cells(Row, i) = Application.WorksheetFunction.VLookup(x, Table2, a, False)
        On Error Resume Next
        a = a + 1
    Next i
        Row = Row + 1
        a = 3
Next x

MsgBox ("Completed Successfully")

End Sub


