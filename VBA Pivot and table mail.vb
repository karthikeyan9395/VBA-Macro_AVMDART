Option Explicit

Sub Codes()

'Application.DisplayAlerts = False
'Call New_Sheet
Application.DisplayAlerts = True
Dim sh As Worksheet
Dim rw As Range
Dim RowCount As Integer
Dim rngFound As Range
RowCount = 0
Dim valueFound As Boolean

Set sh = ActiveSheet
For Each rw In sh.Rows
    If sh.Cells(rw.row, 1).Value = "" Then
      Call SortSomeStuff(RowCount)
      GoTo BREAK
    End If
    RowCount = RowCount + 1
Next rw

BREAK:
    
    Call Codes_1
    Call Macro1_Priority(RowCount)
    Call Macro2_HasBreached
    Call Macro3_WorkNote
    Call Macro4_ClosedToday
    Call Macro5_FollowupToday
   ' Call Mail_small_Text_Outlook
End Sub

Sub SortSomeStuff(row As Integer)
'Remove Duplicates by sorting Ticket number(Asc) and Start Time(Desc)
    Dim val1 As String
    Dim val2 As String
    Dim val3 As String
    val1 = "A2:" & "A" & row
    val2 = "S2:" & "S" & row
    val3 = "A1:" & "Y" & row
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(val1 _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(val2 _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range(val3)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.Range(val3).RemoveDuplicates Columns:=1, Header:= _
        xlYes
End Sub

Sub Codes_1()

Dim sh As Worksheet
Dim rw As Range
Dim RowCount As Long
Dim rngFound As Range
Dim rngFound1 As Range
RowCount = 0
Dim valueFound As Boolean
Set sh = ActiveSheet

For Each rw In sh.Rows
    If sh.Cells(rw.row, 1).Value = "" Then
      GoTo BREAK
    End If
    RowCount = RowCount + 1
    
    If rw.row = 1 Then
       sh.Cells(rw.row, 26).Value = "Actual time left in Hours"
       sh.Cells(rw.row, 27).Value = "To be Closed Today"
       sh.Cells(rw.row, 28).Value = "Work Notes Upto Date"
       sh.Cells(rw.row, 29).Value = "Aging"
       
    Else
    
        sh.Cells(rw.row, 26).Value = sh.Cells(rw.row, 24).Value / 3600
        
        sh.Cells(rw.row, 29).Value = CInt(DateTime.Now - sh.Cells(rw.row, 11).Value)
        
        Dim openPos As Integer
        Dim closePos As Integer
        Dim midBit As String
        Dim wrkDate As Date
        Dim rngLookup As Range
        Dim rngLookup1 As Range
        Dim sDayName As String
        Dim tim As Integer
        
        If sh.Cells(rw.row, 25).Value <> "" Then
            openPos = InStr(sh.Cells(rw.row, 25).Value, "-")
            closePos = InStr(sh.Cells(rw.row, 25).Value, "(")
            wrkDate = Left(sh.Cells(rw.row, 25).Value, 10)
            midBit = Mid(sh.Cells(rw.row, 25).Value, openPos + 2, closePos - openPos - 3)
            On Error Resume Next
            Err.Clear
            sh.Cells(rw.row, 30).Value = Application.WorksheetFunction.VLookup(midBit, rngLookup, 1, False)
            If Err.Number = 0 Then
                'Debug.Print "Found item. The value is " & midBit
                If CInt(DateTime.Now - wrkDate) > 3 Then
                    sh.Cells(rw.row, 28).Value = "No"
                Else
                    sh.Cells(rw.row, 28).Value = "Yes"
                End If
            Else
                Debug.Print "Could not find value: " & midBit
                sh.Cells(rw.row, 28).Value = "No"
            End If
        Else
            sh.Cells(rw.row, 28).Value = "No"
        End If
        
        With ActiveWorkbook.Sheets("Users")
            Set rngLookup = .Range(.Cells(1, 1), .Cells(3, 1).End(xlDown)).Resize(, 3)
            Set rngLookup1 = .Range("B:C")
        End With
        
        If sh.Cells(rw.row, 3).Value = "" Then
            sh.Cells(rw.row, 31).Value = Application.WorksheetFunction.VLookup(sh.Cells(rw.row, 16).Value, rngLookup1, 2, False)
        Else
            sh.Cells(rw.row, 31).Value = Application.WorksheetFunction.VLookup(sh.Cells(rw.row, 3).Value, rngLookup, 3, False)
        End If
            
        sDayName = Format(Date, "dddd")
        
        If sDayName = "Friday" Then
            tim = 72
        Else
            tim = 24
        
        End If
        
        If sh.Cells(rw.row, 26).Value <= tim And sh.Cells(rw.row, 26).Value <> 0 Then
            sh.Cells(rw.row, 27).Value = "Yes"
            Select Case sh.Cells(rw.row, 10).Value
                Case "In Progress"
                    Call focus_Email(sh.Cells(rw.row, 4).Value, sh.Cells(rw.row, 1).Value, sh.Cells(rw.row, 31).Value, CInt(sh.Cells(rw.row, 26).Value), "Closure")
                Case "Acknowledged"
                    Call focus_Email(sh.Cells(rw.row, 4).Value, sh.Cells(rw.row, 1).Value, sh.Cells(rw.row, 31).Value, CInt(sh.Cells(rw.row, 26).Value), "Closure")
                Case "Open"
                    Call focus_Email(sh.Cells(rw.row, 4).Value, sh.Cells(rw.row, 1).Value, sh.Cells(rw.row, 31).Value, CInt(sh.Cells(rw.row, 26).Value), "Closure")
            End Select
        Else
            sh.Cells(rw.row, 27).Value = "No"
        End If
        
        If sh.Cells(rw.row, 21).Value <> "" Then
            Call focus_Email(sh.Cells(rw.row, 4).Value, sh.Cells(rw.row, 1).Value, sh.Cells(rw.row, 31).Value, CInt(sh.Cells(rw.row, 21).Value), "Escalation")
        End If
            
    End If
     
Next rw

BREAK:
'MsgBox (RowCount)

End Sub

Public Function Contains(strBaseString As String, strSearchTerm As String) As Boolean
'Purpose: Returns TRUE if one string exists within another
On Error GoTo ErrorMessage
    Contains = InStr(strBaseString, strSearchTerm)
Exit Function
ErrorMessage:
MsgBox "The database has generated an error. Please contact the database administrator, quoting the following error message: '" & Err.Description & "'", vbCritical, "Database Error"
End
End Function

Sub Macro1_Priority(row As Integer)

'Set up pivot for Priority
Dim wb As Workbook
Dim ws As Worksheet
Dim val1 As String
val1 = "Sheet1!R1C1:R" & row & "C29"

 Sheets.Add(, Sheets(Sheets.Count)).Name = "PivotSheet"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        val1, Version:=6).CreatePivotTable TableDestination:= _
        "PivotSheet!R1C1", TableName:="PivotTable1", DefaultVersion:=6
    

    Sheets("PivotSheet").Select
    Sheets("PivotSheet").Move Before:=Sheets(1)
    'Cells(3, 1).Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Number"), "Count of Number", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Priority")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Group")
        .Orientation = xlRowField
        .Position = 1
    End With
    On Error Resume Next
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Priority").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Priority")
        .PivotItems("3 - Moderate").Visible = False
        .PivotItems("4 - Low").Visible = False
        .PivotItems("5 - Planning").Visible = False
    End With
    On Error Resume Next
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Priority"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable1").TableRange1.Select
    Call pivot_liner
    
    
End Sub

Sub Macro2_HasBreached()
'Set up pivot for HasBreached column value true
 Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("PivotSheet").PivotTables("PivotTable1").PivotCache. _
        CreatePivotTable TableDestination:="PivotSheet!R1C4", TableName:= _
        "PivotTable2", DefaultVersion:=6
    Sheets("PivotSheet").Select
    'Cells(1, 4).Select
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Number"), "Count of Number", xlCount
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Has breached")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Group")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Has breached"). _
        ClearAllFilters
    On Error Resume Next
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Has breached").CurrentPage _
        = "TRUE"
    ActiveSheet.PivotTables("PivotTable2").TableRange1.Select
    Call pivot_liner
    End Sub

Sub Macro3_WorkNote()
'Set up pivot for WorkNote column value "No"
 Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("PivotSheet").PivotTables("PivotTable1").PivotCache. _
        CreatePivotTable TableDestination:="PivotSheet!R1C7", TableName:= _
        "PivotTable3", DefaultVersion:=6
    Sheets("PivotSheet").Select
    'Cells(1, 4).Select

    
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Number"), "Count of Number", xlCount
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Group")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Work Notes Upto Date")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Work Notes Upto Date"). _
        ClearAllFilters
    On Error Resume Next
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Work Notes Upto Date").CurrentPage _
        = "No"
    ActiveSheet.PivotTables("PivotTable3").TableRange1.Select
    Call pivot_liner
End Sub

Sub Macro4_ClosedToday()
'Set up pivot for To be closed tickets column value "Yes" with SLA Running States
 Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("PivotSheet").PivotTables("PivotTable1").PivotCache. _
        CreatePivotTable TableDestination:="PivotSheet!R25C1", TableName:= _
        "PivotTable4", DefaultVersion:=6
    Sheets("PivotSheet").Select
    'Cells(1, 4).Select

    
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Number"), "Count of Number", xlCount
    
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("State")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Group")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("To be Closed Today")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable4").PivotFields("To be Closed Today"). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable4").PivotFields("State"). _
    ClearAllFilters
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("State")
        .PivotItems("Pending Vendor").Visible = False
        .PivotItems("Pending Release").Visible = False
        .PivotItems("Waiting User Reply").Visible = False
        .PivotItems("Scheduled").Visible = False
        .PivotItems("Approval In Progress").Visible = False
    End With
     ActiveSheet.PivotTables("PivotTable4").PivotFields("To be Closed Today").CurrentPage _
        = "Yes"
    ActiveSheet.PivotTables("PivotTable4").TableRange1.Select
    Call pivot_liner
   
End Sub

Sub Macro5_FollowupToday()
'Set up pivot for To be closed tickets column value "Yes" with SLA Pause States
 Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("PivotSheet").PivotTables("PivotTable1").PivotCache. _
        CreatePivotTable TableDestination:="PivotSheet!R25C4", TableName:= _
        "PivotTable5", DefaultVersion:=6
    Sheets("PivotSheet").Select
    'Cells(1, 4).Select

    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Number"), "Count of Number", xlCount
    
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("State")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Group")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("To be Closed Today")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable5").PivotFields("To be Closed Today"). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable5").PivotFields("State"). _
    ClearAllFilters
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("State")
        .PivotItems("In Progress").Visible = False
        .PivotItems("Acknowledged").Visible = False
        .PivotItems("Open").Visible = False
        .PivotItems("").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable5").PivotFields("To be Closed Today").CurrentPage _
        = "Yes"
    ActiveSheet.PivotTables("PivotTable5").TableRange1.Select
    Call pivot_liner
   
End Sub

Sub Mail_small_Text_Outlook()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim strbody1 As String
    Dim strbody2 As String
    Dim strbody3 As String
    Dim strbody4 As String
    Dim rng As Range
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rng4 As Range
    Dim sDayName As String
    Dim subject As String
    
    sDayName = Format(Date, "dddd")
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    If sDayName = "Friday" Then
        subject = "<Action Item> WEEKEND Alert: Watch List -> Priority Tickets to be closed today " & Date
    Else
        subject = "<Action Item> Watch List -> Priority Tickets to be closed today " & Date
    End If
    
    strbody = "Hi All," & "<br>" & "<br>" & _
            "The latest dump of open tickets are placed in below file. Please ensure that we are meeting SLAs, work notes are updated periodically. Notify your onshore counterparts as well." & "<br>" & _
            "<br>" & "https://ch1hub.cognizant.com/sites/SC4562/XLCTeamRepository/Shared%20Documents/02.%20AMS/13.%20Backlog%20Status/AMS_Backlogs_Aug2018_IST.xlsm" & "<br>" & _
            "<br>" & "<br>" & "Tickets that are to be closed today (State - Acknowledged, In Progress, Open). This must be closed today within SLA along with any priority action items that is coming today. " & "<br>" & _
            "<br>" & "<br>"
            
    Set rng = Sheets("PivotSheet").Range("A25:B34").SpecialCells(xlCellTypeVisible)
    
    strbody1 = "<br>" & _
            "Tickets that are to be followed up today (State - Pending Vendor, Pending Release, Approval In Progress, Scheduled, Waiting User Reply). This must be followed up (3 Strike rule for user) today within SLA along with any priority action items that is coming today. " & _
            "<br>" & "<br>" & "<br>"
     
     Set rng1 = Sheets("PivotSheet").Range("D25:E34").SpecialCells(xlCellTypeVisible)
     
     strbody2 = "P1/P2 Tickets that are still open. The P1/P2 ticket(s) should be closed today or downgraded to the correct priority. P1/P2 should not be open more than 2 hours/4 hours respectively." & _
            "<br>" & _
            "<br>" & _
            "<br>"
    
    Set rng2 = Sheets("PivotSheet").Range("A3:B13").SpecialCells(xlCellTypeVisible)
    
     strbody3 = "Work Notes Not updated for more than 2 Business days:  Work Notes must be updated for below open tickets today. The count is not reducing, this count must come down " & _
            "<br>" & _
            "<br>"
     Set rng3 = Sheets("PivotSheet").Range("G3:H13").SpecialCells(xlCellTypeVisible)
           
     strbody4 = "Breached Tickets and still open: Tickets which shows SLA has breached are below. Please find out the reason for this and close this at the earliest." & _
            "<br>" & _
            "<br>"
    Set rng4 = Sheets("PivotSheet").Range("D3:E13").SpecialCells(xlCellTypeVisible)
    
    On Error Resume Next
    With OutMail
        .To = "balamurugan.vijendran@contractor.xlcatlin.com;joshua.m@contractor.xlcatlin.com; daislin.ajit@contractor.xlcatlin.com; harish.chandramohan@contractor.xlcatlin.com; savari.rajendrant@contractor.xlcatlin.com; rajan.muthukrishnan@contractor.xlcatlin.com; divyashanthia.m@contractor.xlcatlin.com; MunilPrakash.G@cognizant.com; Gopinath.Sivagnanam@cognizant.com; Nandan.Purushothaman@cognizant.com; KishanthNelson.SigamoneyNelson@cognizant.com; nirab.sinha@contractor.xlcatlin.com; sivakumar.damodharan@contractor.xlcatlin.com;  VinayKumar.Gudigundla@contractor.xlcatlin.com; sumanjit.raut@contractor.xlcatlin.com; jayasindhan.mohan@contractor.xlcatlin.com; Vikram.Sandhyana@contractor.xlcatlin.com; Narahari.Galla@cognizant.com; charles.chezhyan@contractor.xlcatlin.com; gayathri.balu@contractor.xlcatlin.com; XLCAMSLeads@cognizant.com; Parthiv.Parikh@cognizant.com"
        .CC = "AravindRajan.EP@cognizant.com; Varadarajan.VR@cognizant.com; Srinivasan.Ramachandran2@cognizant.com; Sudhakar.Thirunavukarasu@cognizant.com; Parthiv.Parikh@cognizant.com;Balaji.Selvaraj@cognizant.com;Ganesh.KV@cognizant.com"
        .BCC = ""
        .subject = subject
        .HTMLBody = strbody & RangetoHTML(rng) & "<br>" & "<br>" & _
                    strbody1 & RangetoHTML(rng1) & "<br>" & "<br>" & _
                    strbody2 & RangetoHTML(rng2) & "<br>" & "<br>" & _
                    strbody3 & RangetoHTML(rng3) & "<br>" & "<br>" & _
                    strbody4 & RangetoHTML(rng4) & "<br>" & "<br>"
                    
        '.Attachments.Add ("C:\test.txt")
       ' .Send   'or use
        '.Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function

Function pivot_liner()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Function


Sub focus_Email(email As String, ticketnumber As String, CCemail As String, timeLeft As Integer, scope As String)
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim subject As String
    Dim sDayName As String
    Dim sub1 As String
    Dim sub2 As String
    Dim sub3 As String
    Dim dummyemail As String
    'sub1 = sub2 = sub3 = strbody = subject = sDayName = dummyemail = ""
    sDayName = Format(Date, "dddd")
    
    If email = "" Then
        sub2 = "Unassigned"
        sub3 = " in your resolver queue"
    Else
        sub2 = "Your"
        sub3 = ""
    End If
    
    If sDayName = "Friday" Then
     sub1 = "WEEKEND ALERT:"
    Else
        sub1 = ""
    End If
    
    
    If scope = "Closure" Then
        subject = "<Action Item> " & sub1 & sub2 & " Ticket " & ticketnumber & sub3 & " must be closed today " & Date & " within next " & timeLeft & " hours."
        
    Else
        subject = "<Action Item> " & sub1 & " ESCALATED Ticket " & ticketnumber & sub3 & " must be actioned today " & Date
    End If
    
    strbody = "Hi," & "<br>" & "<br>" & _
            "The ticket " & ticketnumber & " is due for closure  and must be closed within next " & timeLeft & " hours.<br>" & _
            "<br>" & "Please ensure that you are closing this within SLA" & "<br>"
            
    If scope = "Escalation" Then
        'dummyemail = email
        'email = CCemail
        'CCemail = dummyemail
        strbody = "Hi," & "<br>" & "<br>" & _
            "The ticket " & ticketnumber & " is escalated " & timeLeft & " time(s).<br>" & _
            "<br>" & "Please ensure that the prompt action is taken for this ticket" & "<br>"
    End If
    
    If email = "" Then
        email = CCemail
        CCemail = ""
    End If
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    
    On Error Resume Next
    With OutMail
        .To = email
        .CC = CCemail
        .BCC = ""
        .subject = subject
        .HTMLBody = strbody & "<br>" & "<br>"
     '   .Send   'or use
        '.Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



Sub New_Sheet()
'
' Macro1 Macro
'

    Application.DisplayAlerts = False
    Sheets("PivotSheet").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Move Before:=Sheets(1)
    Sheets("Sheet1").Activate
    Sheets("Sheet1").Cells.Select
    Selection.RowHeight = 12.75
    Sheets("Sheet1").Cells(1, 1).Select
End Sub

Sub Waiting_user_reply()

Dim wait As String
wait = "Waiting User Reply"

Worksheets("Sheet1").Activate
Range(Cells(A1)).AutoFilter Field:=10, Criteria1:=wait

End Sub

