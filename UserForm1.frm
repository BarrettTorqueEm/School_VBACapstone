VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Date Chooser"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet

'~~> Prepare the Calendar on Userform
Private Sub UserForm_Initialize()
    '~~> Create a temp sheet for `GenerateCal` to work upon
    Set ws = Sheets.Add
    ws.Visible = xlSheetVeryHidden
    
    GenerateCal Format(Date, "mm/yyyy")
End Sub

'~~> Next Month
Private Sub CommandButton43_Click()
    Dim dat As Date
    dat = DateAdd("m", -1, DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), 1))
    GenerateCal Format(dat, "mm/yyyy")
    CommandButton45.Caption = Format(dat, "mmm - yyyy")
End Sub

'~~> Previous Month
Private Sub CommandButton44_Click()
    Dim dat As Date
    dat = DateAdd("m", 1, DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), Val(Format(CommandButton45.Caption, "MM")), 1))
    GenerateCal Format(dat, "mm/yyyy")
    CommandButton45.Caption = Format(dat, "mmm - yyyy")
End Sub

'~~> Ok Button
Private Sub CommandButton53_Click()
    MsgBox TextBox1.Text
End Sub

'~~> Cancel Button
Private Sub CommandButton54_Click()
    Unload Me
End Sub

'~~> Generate Sheet
'~~> Code based on http://support.microsoft.com/kb/150774
Private Sub GenerateCal(dt As String)
    With ws
        .Cells.Clear
        StartDay = DateValue(dt)
        ' Check if valid date but not the first of the month
        ' -- if so, reset StartDay to first day of month.
        If Day(StartDay) <> 1 Then
            StartDay = DateValue(Month(StartDay) & "/1/" & _
                Year(StartDay))
        End If
        ' Prepare cell for Month and Year as fully spelled out.
        .Range("a1").NumberFormat = "mmmm yyyy"
        ' Center the Month and Year label across a1:g1 with appropriate
        ' size, height and bolding.
        With .Range("a1:g1")
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Font.Size = 18
            .Font.Bold = True
            .RowHeight = 35
        End With
        ' Prepare a2:g2 for day of week labels with centering, size,
        ' height and bolding.
        With .Range("a2:g2")
            .ColumnWidth = 11
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = xlHorizontal
            .Font.Size = 12
            .Font.Bold = True
            .RowHeight = 20
        End With
        ' Put days of week in a2:g2.
        .Range("a2") = "Sunday"
        .Range("b2") = "Monday"
        .Range("c2") = "Tuesday"
        .Range("d2") = "Wednesday"
        .Range("e2") = "Thursday"
        .Range("f2") = "Friday"
        .Range("g2") = "Saturday"
        ' Prepare a3:g7 for dates with left/top alignment, size, height
        ' and bolding.
        With .Range("a3:g8")
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlTop
            .Font.Size = 18
            .Font.Bold = True
            .RowHeight = 21
        End With
        ' Put inputted month and year fully spelling out into "a1".
        .Range("a1").Value = Application.Text(dt, "mmmm yyyy")
        ' Set variable and get which day of the week the month starts.
        DayofWeek = Weekday(StartDay)
        ' Set variables to identify the year and month as separate
        ' variables.
        CurYear = Year(StartDay)
        CurMonth = Month(StartDay)
        ' Set variable and calculate the first day of the next month.
        FinalDay = DateSerial(CurYear, CurMonth + 1, 1)
        ' Place a "1" in cell position of the first day of the chosen
        ' month based on DayofWeek.
        Select Case DayofWeek
            Case 1
                .Range("a3").Value = 1
            Case 2
                .Range("b3").Value = 1
            Case 3
                .Range("c3").Value = 1
            Case 4
                .Range("d3").Value = 1
            Case 5
                .Range("e3").Value = 1
            Case 6
                .Range("f3").Value = 1
            Case 7
                .Range("g3").Value = 1
        End Select
        ' Loop through .Range a3:g8 incrementing each cell after the "1"
        ' cell.
        For Each cell In .Range("a3:g8")
            RowCell = cell.Row
            ColCell = cell.Column
            ' Do if "1" is in first column.
            If cell.Column = 1 And cell.Row = 3 Then
            ' Do if current cell is not in 1st column.
            ElseIf cell.Column <> 1 Then
                If cell.Offset(0, -1).Value >= 1 Then
                    cell.Value = cell.Offset(0, -1).Value + 1
                    ' Stop when the last day of the month has been
                    ' entered.
                    If cell.Value > (FinalDay - StartDay) Then
                        cell.Value = ""
                        ' Exit loop when calendar has correct number of
                        ' days shown.
                        Exit For
                    End If
                End If
            ' Do only if current cell is not in Row 3 and is in Column 1.
            ElseIf cell.Row > 3 And cell.Column = 1 Then
                cell.Value = cell.Offset(-1, 6).Value + 1
                ' Stop when the last day of the month has been entered.
                If cell.Value > (FinalDay - StartDay) Then
                    cell.Value = ""
                    ' Exit loop when calendar has correct number of days
                    ' shown.
                    Exit For
                End If
            End If
        Next
    
        ' Create Entry cells, format them centered, wrap text, and border
        ' around days.
        For x = 0 To 5
            .Range("A4").Offset(x * 2, 0).EntireRow.Insert
            With .Range("A4:G4").Offset(x * 2, 0)
                .RowHeight = 65
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
                .WrapText = True
                .Font.Size = 10
                .Font.Bold = False
                ' Unlock these cells to be able to enter text later after
                ' sheet is protected.
                .Locked = False
            End With
            ' Put border around the block of dates.
            With .Range("A3").Offset(x * 2, 0).Resize(2, _
            7).Borders(xlLeft)
                .Weight = xlThick
                .ColorIndex = xlAutomatic
            End With
    
            With .Range("A3").Offset(x * 2, 0).Resize(2, _
            7).Borders(xlRight)
                .Weight = xlThick
                .ColorIndex = xlAutomatic
            End With
            .Range("A3").Offset(x * 2, 0).Resize(2, 7).BorderAround _
               Weight:=xlThick, ColorIndex:=xlAutomatic
        Next
        If .Range("A13").Value = "" Then .Range("A13").Offset(0, 0) _
           .Resize(2, 8).EntireRow.Delete
    
        ' Resize window to show all of calendar (may have to be adjusted
        ' Allow screen to redraw with calendar showing.
        Application.ScreenUpdating = True
        
        '~~> Update Dates on command button
        CommandButton1.Caption = .Range("A3").Text
        CommandButton2.Caption = .Range("B3").Text
        CommandButton3.Caption = .Range("C3").Text
        CommandButton4.Caption = .Range("D3").Text
        CommandButton5.Caption = .Range("E3").Text
        CommandButton6.Caption = .Range("F3").Text
        CommandButton7.Caption = .Range("G3").Text
        
        CommandButton8.Caption = .Range("A5").Text
        CommandButton9.Caption = .Range("B5").Text
        CommandButton10.Caption = .Range("C5").Text
        CommandButton11.Caption = .Range("D5").Text
        CommandButton12.Caption = .Range("E5").Text
        CommandButton13.Caption = .Range("F5").Text
        CommandButton14.Caption = .Range("G5").Text
        
        CommandButton15.Caption = .Range("A7").Text
        CommandButton16.Caption = .Range("B7").Text
        CommandButton17.Caption = .Range("C7").Text
        CommandButton18.Caption = .Range("D7").Text
        CommandButton19.Caption = .Range("E7").Text
        CommandButton20.Caption = .Range("F7").Text
        CommandButton21.Caption = .Range("G7").Text
        
        CommandButton22.Caption = .Range("A9").Text
        CommandButton23.Caption = .Range("B9").Text
        CommandButton24.Caption = .Range("C9").Text
        CommandButton25.Caption = .Range("D9").Text
        CommandButton26.Caption = .Range("E9").Text
        CommandButton27.Caption = .Range("F9").Text
        CommandButton28.Caption = .Range("G9").Text
        
        CommandButton29.Caption = .Range("A11").Text
        CommandButton30.Caption = .Range("B11").Text
        CommandButton31.Caption = .Range("C11").Text
        CommandButton32.Caption = .Range("D11").Text
        CommandButton33.Caption = .Range("E11").Text
        CommandButton34.Caption = .Range("F11").Text
        CommandButton35.Caption = .Range("G11").Text
        
        CommandButton46.Caption = .Range("A13").Text
        CommandButton47.Caption = .Range("B13").Text
        CommandButton48.Caption = .Range("C13").Text
        CommandButton49.Caption = .Range("D13").Text
        CommandButton50.Caption = .Range("E13").Text
        CommandButton51.Caption = .Range("F13").Text
        CommandButton52.Caption = .Range("G13").Text
    End With
End Sub

'~~> Delete the Temp Sheet that was created
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

'~~> This section simply updates the date in the text box when a button is pressed
Private Sub CommandButton1_Click()
    If CommandButton1.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton1.Caption)
End Sub
Private Sub CommandButton2_Click()
    If CommandButton2.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton2.Caption)
End Sub
Private Sub CommandButton3_Click()
    If CommandButton3.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton3.Caption)
End Sub
Private Sub CommandButton4_Click()
    If CommandButton4.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton4.Caption)
End Sub
Private Sub CommandButton5_Click()
    If CommandButton5.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton5.Caption)
End Sub
Private Sub CommandButton6_Click()
    If CommandButton6.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton6.Caption)
End Sub
Private Sub CommandButton7_Click()
    If CommandButton7.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton7.Caption)
End Sub
Private Sub CommandButton8_Click()
    If CommandButton8.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton8.Caption)
End Sub
Private Sub CommandButton9_Click()
    If CommandButton9.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton9.Caption)
End Sub
Private Sub CommandButton10_Click()
    If CommandButton10.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton10.Caption)
End Sub
Private Sub CommandButton11_Click()
    If CommandButton11.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton11.Caption)
End Sub
Private Sub CommandButton12_Click()
    If CommandButton12.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton12.Caption)
End Sub
Private Sub CommandButton13_Click()
    If CommandButton13.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton13.Caption)
End Sub
Private Sub CommandButton14_Click()
    If CommandButton14.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton14.Caption)
End Sub
Private Sub CommandButton15_Click()
    If CommandButton15.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton15.Caption)
End Sub
Private Sub CommandButton16_Click()
    If CommandButton16.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton16.Caption)
End Sub
Private Sub CommandButton17_Click()
    If CommandButton17.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton17.Caption)
End Sub
Private Sub CommandButton18_Click()
    If CommandButton18.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton18.Caption)
End Sub
Private Sub CommandButton19_Click()
    If CommandButton19.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton19.Caption)
End Sub
Private Sub CommandButton20_Click()
    If CommandButton20.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton20.Caption)
End Sub
Private Sub CommandButton21_Click()
    If CommandButton21.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton21.Caption)
End Sub
Private Sub CommandButton22_Click()
    If CommandButton22.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton22.Caption)
End Sub
Private Sub CommandButton23_Click()
    If CommandButton23.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton23.Caption)
End Sub
Private Sub CommandButton24_Click()
    If CommandButton24.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton24.Caption)
End Sub
Private Sub CommandButton25_Click()
    If CommandButton25.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton25.Caption)
End Sub
Private Sub CommandButton26_Click()
    If CommandButton26.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton26.Caption)
End Sub
Private Sub CommandButton27_Click()
    If CommandButton27.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton27.Caption)
End Sub
Private Sub CommandButton28_Click()
    If CommandButton28.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton28.Caption)
End Sub
Private Sub CommandButton29_Click()
    If CommandButton29.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton29.Caption)
End Sub
Private Sub CommandButton30_Click()
    If CommandButton30.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton30.Caption)
End Sub
Private Sub CommandButton31_Click()
    If CommandButton31.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton31.Caption)
End Sub
Private Sub CommandButton32_Click()
    If CommandButton32.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton32.Caption)
End Sub
Private Sub CommandButton33_Click()
    If CommandButton33.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton33.Caption)
End Sub
Private Sub CommandButton34_Click()
    If CommandButton34.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton34.Caption)
End Sub
Private Sub CommandButton35_Click()
    If CommandButton35.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton35.Caption)
End Sub
Private Sub CommandButton46_Click()
    If CommandButton46.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton46.Caption)
End Sub
Private Sub CommandButton47_Click()
    If CommandButton47.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton47.Caption)
End Sub
Private Sub CommandButton48_Click()
    If CommandButton48.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton48.Caption)
End Sub
Private Sub CommandButton49_Click()
    If CommandButton49.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton49.Caption)
End Sub
Private Sub CommandButton50_Click()
    If CommandButton50.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton50.Caption)
End Sub
Private Sub CommandButton51_Click()
    If CommandButton51.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton51.Caption)
End Sub
Private Sub CommandButton52_Click()
    If CommandButton52.Caption <> "" Then TextBox1.Text = DateSerial(Val(Format(CommandButton45.Caption, "YYYY")), _
    Val(Format(CommandButton45.Caption, "MM")), CommandButton52.Caption)
End Sub

