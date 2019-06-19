'**********************************************************************************************************************'
' Programmer ID : 563525                        **************  @@@@@   @  @@ *************************************    '
' Author        : Karthikeyan Selvan            **************  @       @ @  *********************************         '
' Program       : AVM Dart bulk upload         *************** @@@@@   @@   ****************************              '
' Date          : 03/31/2018 13:00 hrs          ***************     @   @ @  *********************************         '
' Efforts       : 24 Hrs                        *************** @@@@@ # @  @@ *************************************    '
'**********************************************************************************************************************'

Dim dbsheet1 As Worksheet
Dim dbsheet3 As Worksheet
Dim dbsheet4 As Worksheet

Sub Main()

Call delsheets 'Delete data from Original and Sheet4

Call CopyIncident 'Copy SNOW Incidents workbook to current workbook

Call CopyItasks 'Copy SNOW Service requests workbook to current workbook

Call Generate 'Does all the formatting for AVM Dart

Call Rename 'Takes a backup of previous report for future purpose

Call CopyFinal 'Copy the final formatted data to USIns_Dart.xlsx

End Sub

Sub delsheets()

Set dbsheet4 = ActiveWorkbook.Sheets("Sheet4")
Set dbsheet1 = ActiveWorkbook.Sheets("Original")

Application.DisplayAlerts = False

dbsheet4.delete
dbsheet1.delete

Worksheets.Add.Name = "Sheet4"
Worksheets.Add.Name = "Original"

End Sub

Sub CopyIncident()

Dim INC As Workbook 'Incdient SNOW

Set dbsheet1 = ActiveWorkbook.Sheets("Original")

Set INC = Workbooks.Open("C:\Users\User\Desktop\AVMDART\Incident.xlsx")
Set incsheet = ActiveWorkbook.Sheets("Page 1")
r1 = incsheet.Cells(Rows.Count, 1).End(xlUp).row

INC.Activate

Application.CutCopyMode = False

incsheet.Range(Cells(1, 1), Cells(r1, 10)).Copy

dbsheet1.Activate

dbsheet1.Range(Cells(1, 1), Cells(r1, 10)).PasteSpecial

Application.CutCopyMode = False

INC.Close

End Sub

Sub CopyItasks()

Dim ITASK As Workbook 'Service Requests SNOW

Set dbsheet1 = ActiveWorkbook.Sheets("Original")

Set ITASK = Workbooks.Open("C:\Users\User\Desktop\AVMDART\Service Request.xlsx")
Set itasksheet = ActiveWorkbook.Sheets("Page 1")

r1 = itasksheet.Cells(Rows.Count, 1).End(xlUp).row
r2 = dbsheet1.Cells(Rows.Count, 1).End(xlUp).row

ITASK.Activate

Application.CutCopyMode = False

itasksheet.Range(Cells(2, 1), Cells(r1, 10)).Copy

dbsheet1.Activate

dbsheet1.Range(Cells(r2 + 1, 1), Cells(r1 + r2, 10)).PasteSpecial

Application.CutCopyMode = False

ITASK.Close

End Sub

Sub Generate()

Set dbsheet1 = ActiveWorkbook.Sheets("Original")
Set dbsheet3 = ActiveWorkbook.Sheets("Copybook")
Set dbsheet4 = ActiveWorkbook.Sheets("Sheet4")

 r1 = dbsheet1.Cells(Rows.Count, 1).End(xlUp).row
 r3 = dbsheet3.Cells(Rows.Count, 1).End(xlUp).row
' r4 = dbsheet4.Cells(Rows.Count, 1).End(xlUp).Row
 
 'dbsheet1.Range(Cells(2, 1), Cells(r1, 1)).Select
 'dbsheet4.Range(Cells(2, 1), Cells(r1, 1)).Select
 
  Worksheets("Sheet4").Activate
    dbsheet4.Range("A1") = "Number"
    dbsheet4.Range("B1") = "Priority"
    dbsheet4.Range("C1") = "Open Date/Time"
    dbsheet4.Range("D1") = "Assignee"
    dbsheet4.Range("E1") = "Ticket type"
    dbsheet4.Range("F1") = "Group"
    dbsheet4.Range("G1") = "Actual Start Time"
    dbsheet4.Range("H1") = "Actual Close Time"
    dbsheet4.Range("I1") = "Closed Date"
    dbsheet4.Range("J1") = "Ticket Status"
    dbsheet4.Range("K1") = "Resolved by"
    dbsheet4.Range("L1") = "Severity"
    dbsheet4.Range("M1") = "Applications"
    dbsheet4.Range("N1") = "Client User ID"
    dbsheet4.Range("O1") = "TicketSource"
    dbsheet4.Range("P1") = "KEDB Updated"
    dbsheet4.Range("Q1") = "Met Response"
    dbsheet4.Range("R1") = "Met Ack"
    dbsheet4.Range("S1") = "Met Resolution SLA"
 
 
Worksheets("Original").Activate
dbsheet1.Range(Cells(2, 1), Cells(r1, 1)).Copy Destination:=dbsheet4.Cells(2, 1) 'Incident
dbsheet1.Range(Cells(2, 2), Cells(r1, 2)).Copy Destination:=dbsheet4.Cells(2, 2) 'Priority
dbsheet1.Range(Cells(2, 3), Cells(r1, 3)).Copy Destination:=dbsheet4.Cells(2, 3) 'Open date

Call RetrieveID 'Vlookup to get CTSID wrt to Assignee name
dbsheet1.Range(Cells(2, 11), Cells(r1, 11)).Copy Destination:=dbsheet4.Cells(2, 4) 'Assignee

dbsheet1.Range(Cells(2, 5), Cells(r1, 5)).Copy Destination:=dbsheet4.Cells(2, 5) 'Ticket type
dbsheet1.Range(Cells(2, 8), Cells(r1, 8)).Copy Destination:=dbsheet4.Cells(2, 6) 'Group
dbsheet1.Range(Cells(2, 3), Cells(r1, 3)).Copy Destination:=dbsheet4.Cells(2, 7) 'Actual Start time
dbsheet1.Range(Cells(2, 4), Cells(r1, 4)).Copy Destination:=dbsheet4.Cells(2, 8) 'Actual Close time

Call Dteadd 'Add 1 min to close time
Worksheets("Original").Activate
dbsheet1.Range(Cells(2, 6), Cells(r1, 6)).Copy Destination:=dbsheet4.Cells(2, 10) 'Ticket Status
dbsheet1.Range(Cells(2, 2), Cells(r1, 2)).Copy Destination:=dbsheet4.Cells(2, 12) 'Severity

Call AppName 'Search for a string in Config Item and replaces with key app names(Hardcoded)
Call Applications 'Vlookup to get application names with key app names
dbsheet1.Range(Cells(2, 13), Cells(r1, 13)).Copy Destination:=dbsheet4.Cells(2, 13) 'Applications



Worksheets("Sheet4").Activate
dbsheet4.Range(Cells(2, 4), Cells(r1, 4)).Copy Destination:=dbsheet4.Cells(2, 11) 'Resolved by
dbsheet4.Range(Cells(2, 4), Cells(r1, 4)).Copy Destination:=dbsheet4.Cells(2, 14) 'Client User ID
dbsheet4.Range(Cells(2, 15), Cells(r1, 15)) = "Ticketing Tool" 'TicketSource
dbsheet4.Range(Cells(2, 16), Cells(r1, 16)) = "Added" 'KEDB Updated
dbsheet4.Range(Cells(2, 17), Cells(r1, 17)) = "Yes" 'Met Response
dbsheet4.Range(Cells(2, 18), Cells(r1, 18)) = "Yes" 'Met Ack
dbsheet4.Range(Cells(2, 19), Cells(r1, 19)) = "Yes" 'Met Resolution SLA

End Sub


Sub RetrieveID()
On Error Resume Next

Dim row As Integer
Dim clm As Integer

row = 2 'Second row
clm = 11 ' Column K of Original

Set dbsheet1 = ActiveWorkbook.Sheets("Original")
Set dbsheet3 = ActiveWorkbook.Sheets("Copybook")

 r1 = dbsheet1.Cells(Rows.Count, 1).End(xlUp).row
 r3 = dbsheet3.Cells(Rows.Count, 1).End(xlUp).row
' r4 = dbsheet4.Cells(Rows.Count, 1).End(xlUp).Row

Worksheets("Original").Activate
Table1 = dbsheet1.Range(Cells(2, 7), Cells(r1, 7)) 'Name Column from Original
Worksheets("Copybook").Activate
Table2 = dbsheet3.Range(Cells(2, 1), Cells(r3, 2)) ' Range of Table 2 for Vlookup

Worksheets("Original").Activate
For Each x In Table1
  dbsheet1.Cells(row, clm) = Application.WorksheetFunction.VLookup(x, Table2, 2, False)
  row = row + 1
Next x

End Sub

Sub AppName()

Dim row As Integer
Dim clm As Integer

row = 2 'Second row
clm = 12 ' Column L of Original

Set dbsheet1 = ActiveWorkbook.Sheets("Original")
Set dbsheet3 = ActiveWorkbook.Sheets("Copybook")

 r1 = dbsheet1.Cells(Rows.Count, 1).End(xlUp).row
 r3 = dbsheet3.Cells(Rows.Count, 4).End(xlUp).row
' r4 = dbsheet4.Cells(Rows.Count, 1).End(xlUp).Row

Worksheets("Original").Activate
Table1 = dbsheet1.Range(Cells(2, 10), Cells(r1, 10)) 'Config_Item Column from Original


For Each x In Table1

    If InStr(1, x, "GAPS") <> 0 Then
       dbsheet1.Cells(row, clm) = "GAPS"
            
            ElseIf InStr(1, x, "WINS") <> 0 Then
                    dbsheet1.Cells(row, clm) = "WINS"
                
            ElseIf InStr(1, x, "SEDLAK") <> 0 Then
                    dbsheet1.Cells(row, clm) = "SEDLAK"
                
            ElseIf InStr(1, x, "WebApp") <> 0 Then
                    dbsheet1.Cells(row, clm) = "WebApp"
                 
            ElseIf InStr(1, x, "iCAN") <> 0 Then
                    dbsheet1.Cells(row, clm) = "ICAN"
                 
            ElseIf InStr(1, x, "WIS") <> 0 Then
                    dbsheet1.Cells(row, clm) = "WIS"
                 
            ElseIf InStr(1, x, "UWS") <> 0 Then
                    dbsheet1.Cells(row, clm) = "UWS"
                
            Else
                    dbsheet1.Cells(row, clm) = "Unknown"
                    Call unknown(row, clm)
                    
    End If

row = row + 1
Next x

End Sub

Sub unknown(ByVal row As Integer, ByVal clm As Integer)

k = dbsheet1.Cells(row, 9).Value

            If InStr(1, k, "GAPS") <> 0 Then
               dbsheet1.Cells(row, clm) = "GAPS"
                    
                    ElseIf InStr(1, k, "WINS") <> 0 Then
                            dbsheet1.Cells(row, clm) = "WINS"
                        
                    ElseIf InStr(1, k, "SEDLAK") <> 0 Then
                            dbsheet1.Cells(row, clm) = "SEDLAK"
                        
                    ElseIf InStr(1, k, "WebApp") <> 0 Then
                            dbsheet1.Cells(row, clm) = "WebApp"
                         
                    ElseIf InStr(1, k, "iCan") <> 0 Then
                            dbsheet1.Cells(row, clm) = "ICAN"
                         
                    ElseIf InStr(1, k, "WIS") <> 0 Then
                            dbsheet1.Cells(row, clm) = "WIS"
                         
                    ElseIf InStr(1, k, "UWS") <> 0 Then
                            dbsheet1.Cells(row, clm) = "UWS"
                            
                    ElseIf InStr(1, k, "MyAnalysis") <> 0 Then
                            dbsheet1.Cells(row, clm) = "GAPS"
                            
                    ElseIf InStr(1, k, "Mayfare") <> 0 Then
                            dbsheet1.Cells(row, clm) = "SEDLAK"
                            
                    ElseIf InStr(1, k, "gaps") <> 0 Then
                            dbsheet1.Cells(row, clm) = "GAPS"
                    Else
                            dbsheet1.Cells(row, clm) = "Unknown"
            End If
            
End Sub


Sub Applications()

Dim row As Integer
Dim clm As Integer

row = 2 'Second row
clm = 13 'Column M of Original

Set dbsheet1 = ActiveWorkbook.Sheets("Original")
Set dbsheet3 = ActiveWorkbook.Sheets("Copybook")

 r1 = dbsheet1.Cells(Rows.Count, 10).End(xlUp).row
 r2 = dbsheet3.Cells(Rows.Count, 3).End(xlUp).row
 r3 = dbsheet3.Cells(Rows.Count, 4).End(xlUp).row

Worksheets("Original").Activate
Table1 = dbsheet1.Range(Cells(2, 12), Cells(r1, 12)) 'Config_Item Column from Original

Worksheets("Copybook").Activate

Table2 = dbsheet3.Range(Cells(2, 2), Cells(r2, 3)) 'Default Primary applications
Table3 = dbsheet3.Range(Cells(2, 4), Cells(r3, 5)) 'Standard application from Config Item

Worksheets("Original").Activate

For Each i In Table1
    
    If dbsheet1.Cells(row, 12) = "Unknown" Then
    
        temp = dbsheet1.Cells(row, 11).Value
        dbsheet1.Cells(row, clm) = Application.WorksheetFunction.VLookup(temp, Table2, 2, False)
    
    Else
        dbsheet1.Cells(row, clm) = Application.WorksheetFunction.VLookup(i, Table3, 2, False)
    End If
  
  row = row + 1
Next i

End Sub

Sub Dteadd()

Dim Mdate As Date
Dim row As Integer
Dim clm As Integer

row = 2 'Second row
clm = 9 'Column I of Original

Set dbsheet4 = ActiveWorkbook.Sheets("Sheet4")
Worksheets("Sheet4").Activate

 r4 = dbsheet4.Cells(Rows.Count, 8).End(xlUp).row
Range(Cells(2, 9), Cells(r4, 9)).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"

For i = 2 To r4

dbsheet4.Cells(row, clm) = DateAdd("n", 1, (Cells(i, 8)))
row = row + 1

Next i

'Range(Cells(2, 9), Cells(r4, 9)).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"

End Sub



Sub Rename()
Dim Path As String
Dim Dtmod As String
Dim s As String
Dim Newname As String
Dim Oldname As String

Path = Dir("C:\Users\User\Desktop\AVMDART\USIns_Dart.xlsx")

If Path <> "" Then

    Dtmod = FileDateTime("C:\Users\User\Desktop\AVMDART\USIns_Dart.xlsx") 'Fetches last date modified
    
    s = Format(Dtmod, "yyyymmdd_hhmmss") 'Stores in s as 20180331_235959
    
    temp = "C:\Users\User\Desktop\AVMDART\" 'Path to AVM Dart
    
    Oldname = "USIns_Dart.xlsx"
    Newname = Oldname & s & ".xlsx" 'Newname will be USIns_Dart20180331_235959
    
    Name temp & Oldname As temp & Newname 'Renaming procedure

End If

Workbooks.Add
ActiveWorkbook.SaveAs Filename:="C:\Users\User\Desktop\AVMDART\USIns_Dart.xlsx"
ActiveWorkbook.Close

End Sub


Sub CopyFinal()
Dim wbtarget As Workbook 'workbook where the data is to be pasted
Dim wbThis As Worksheet 'workbook from where the data is to be copied
Dim temp As Worksheet

Set dbsheet4 = ActiveWorkbook.Sheets("Sheet4")
r4 = dbsheet4.Cells(Rows.Count, 1).End(xlUp).row

Set wbThis = ActiveWorkbook.Sheets("Sheet4")

wbThis.Activate

Application.CutCopyMode = False

wbThis.Range(Cells(1, 1), Cells(r4, 19)).Copy

Set wbtarget = Workbooks.Open("C:\Users\User\Desktop\AVMDART\USIns_Dart.xlsx")

wbtarget.Activate

Set temp = ActiveWorkbook.Sheets("Sheet1")

'temp.Range("A1").Select

temp.Range(Cells(1, 1), Cells(r4, 19)).PasteSpecial

'temp.PasteSpecial

Application.CutCopyMode = False

wbtarget.Save

wbtarget.Close

ActiveWorkbook.Save
ActiveWorkbook.Close

Application.Quit

End Sub
