Sub Increase_and_Send()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim ControlPanel As ThisWorkbook
    Dim Spin As Workbook
    Dim PriceFile As Workbook
    
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    PreviousMonthName = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    
    Set ControlPanel = ThisWorkbook
    Estimator = ControlPanel.Sheets("Paths").Range("B2").Value
    Set Spin = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Spin File\" & FutureMonthNameDir & "\SPIN.xlsx", ReadOnly:=True)
    ActiveSheet.Name = "SPINDATA"
    LastRowSpin = Spin.Sheets("SPINDATA").Range("A" & Rows.Count).End(xlUp).Row
    
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterNextMonth, Operator:=xlFilterDynamic
        If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = FutureMonthName & " Increases"
                Spin.Sheets("Spindata").Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
                Spin.Sheets(FutureMonthName & " Increases").Range("A1").PasteSpecial Paste:=xlPasteValues
                Spin.Sheets(FutureMonthName & " Increases").Range("A1").PasteSpecial Paste:=xlFormats
                Application.CutCopyMode = False
                Spin.Sheets(1).AutoFilter.ShowAllData
            Else
        Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterThisMonth, Operator:=xlFilterDynamic
    If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = ActiveMonthName & " Increases"
            Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
            Spin.Sheets(ActiveMonthName & " Increases").Range("A1").PasteSpecial Paste:=xlPasteValues
            Spin.Sheets(ActiveMonthName & " Increases").Range("A1").PasteSpecial Paste:=xlFormats
            Application.CutCopyMode = False
            Spin.Sheets(1).AutoFilter.ShowAllData
        Else
    Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterLastMonth, Operator:=xlFilterDynamic
    If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = PreviousMonthName & " Increases"
            Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
            Spin.Sheets(PreviousMonthName & " Increases").Range("A1").PasteSpecial Paste:=xlPasteValues
            Spin.Sheets(PreviousMonthName & " Increases").Range("A1").PasteSpecial Paste:=xlFormats
            Application.CutCopyMode = False
            Spin.Sheets(1).AutoFilter.ShowAllData
        Else
    Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    FinalAccount = ControlPanel.Sheets("Paths").Range("D" & Rows.Count).End(xlUp).Row
    Dim Start As Integer
    For Start = 3 To FinalAccount
        AccountPath = ControlPanel.Sheets("Paths").Range("D" & Start).Value
        AccountName = ControlPanel.Sheets("Paths").Range("A" & Start).Value
        Set PriceFile = Workbooks.Open(AccountPath, ReadOnly:=True)
        LastRowPriceFile = PriceFile.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
        Spin.Sheets.Add.Name = "Temp"
        Spin.Sheets("Temp").Range("A1").Value = "HOS Code"
        Spin.Sheets("Temp").Range("B1").Value = "Supplier Name"
        Spin.Sheets("Temp").Range("C1").Value = "Spin Comment"
        Spin.Sheets("Temp").Range("D1").Value = "Average Increase"
        Spin.Sheets("Temp").Range("E1").Value = "Due Date"
        Spin.Sheets("Temp").Range("F1").Value = "No. Products"
        PriceFile.Sheets("Price File").Range("BJ12:BJ" & LastRowPriceFile).Copy
        Spin.Sheets("Temp").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Spin.Sheets("Temp").Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
        lastRowTemp = Spin.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Spin.Sheets("Temp").Range("B2:B" & lastRowTemp).Formula = "=VLOOKUP(A2,SPINDATA!C:D,2,FALSE)"
        Spin.Sheets("Temp").Range("C2:C" & lastRowTemp).Formula = "=IFERROR(IFERROR(IFERROR(VLOOKUP(A2,'" & FutureMonthName & " Increases'!C:G,3,FALSE),VLOOKUP(A2,'" & ActiveMonthName & " Increases'!C:G,3,FALSE)),VLOOKUP(A2,'" & PreviousMonthName & " Increases'!C:G,3,FALSE)),""No Upcoming Increases"")"
        Spin.Sheets("Temp").Range("D2:D" & lastRowTemp).Formula = "=IFERROR(IFERROR(IFERROR(VLOOKUP(A2,'" & FutureMonthName & " Increases'!C:G,4,FALSE),VLOOKUP(A2,'" & ActiveMonthName & " Increases'!C:G,4,FALSE)),VLOOKUP(A2,'" & PreviousMonthName & " Increases'!C:G,4,FALSE)),""No Upcoming Increases"")"
        Spin.Sheets("Temp").Range("E2:E" & lastRowTemp).Formula = "=IFERROR(IFERROR(IFERROR(VLOOKUP(A2,'" & FutureMonthName & " Increases'!C:G,5,FALSE),VLOOKUP(A2,'" & ActiveMonthName & " Increases'!C:G,5,FALSE)),VLOOKUP(A2,'" & PreviousMonthName & " Increases'!C:G,5,FALSE)),""No Upcoming Increases"")"
        Spin.Sheets("Temp").Range("F2:F" & lastRowTemp).Formula = "=COUNTIFS('[" & PriceFile.Name & "]Price File'!$BJ:$BJ,A2)"
        Spin.Sheets("Temp").Range("A1:F" & lastRowTemp).Copy
        Spin.Sheets("Temp").Range("A1:F" & lastRowTemp).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Spin.Sheets("Temp").Range("D1:D" & lastRowTemp).NumberFormat = "0.00%"
        Spin.Sheets("Temp").Range("E1:E" & lastRowTemp).NumberFormat = "dd/mm/yyyy"
        Spin.Sheets("Temp").Range("A1:E" & lastRowTemp).AutoFilter Field:=3, Criteria1:="No Upcoming Increases"
        If Spin.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Spin.Sheets("Temp").AutoFilter.ShowAllData
        End If
        Spin.Sheets("Temp").AutoFilter.ShowAllData
        lastRowTemp = Spin.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Spin.Sheets("Temp").Range("A1:F1").Interior.Color = RGB(180, 198, 231)
        Spin.Sheets("Temp").Range("A2:F" & lastRowTemp).Interior.Color = RGB(219, 219, 219)
        Spin.Sheets("Temp").Range("A1:F1").Font.Bold = True
        Spin.Sheets("Temp").Columns.AutoFit
        Spin.Sheets("Temp").Rows.AutoFit
        Spin.Sheets("Temp").Range("A1:F" & lastRowTemp).HorizontalAlignment = xlCenter
        Spin.Sheets("Temp").Range("A1:F" & lastRowTemp).VerticalAlignment = xlCenter
        Spin.Sheets("Temp").Range("A1:F" & lastRowTemp).Borders.LineStyle = xlContinuous
        Spin.Sheets("Temp").Range("A1:F" & lastRowTemp).Sort Key1:=Spin.Sheets("temp").Range("E1"), Order1:=xlDescending, Header:=xlYes
        Spin.Sheets("Temp").Copy
        ActiveWorkbook.SaveAs Estimator & "\Spin Notification\" & AccountName & " - SLP Changes.xlsx", FileFormat:=51
        ActiveWorkbook.Close True
        Spin.Sheets("Temp").Delete
        PriceFile.Close False
    Next Start
    Spin.Close False
    
    answer = MsgBox("Do you want to send emails automatically?", vbQuestion + vbYesNo + vbDefaultButton2, "Attention")
    If answer = vbYes Then
    Dim MyOutlook As Object
    Set MyOutlook = CreateObject("Outlook.Application")
    For Start = 3 To FinalAccount
        If ControlPanel.Sheets("Paths").Range("F" & Start).Value <> "NO" Then
            Dim MyMail As Object
            Set MyMail = MyOutlook.CreateItem(olMailItem)
            AccountMngName = ControlPanel.Sheets("Paths").Range("G" & Start).Value
            Signature = MyMail.Body
            MyMail.To = ControlPanel.Sheets("Paths").Range("F" & Start).Value
            MyMail.Subject = ControlPanel.Sheets("Paths").Range("A" & Start).Value & " SLP Changes (" & FutureMonthName & " 2022)"
            MyMail.Body = "Hi " & AccountMngName & ", Please see attached possible upcoming increases for " & ControlPanel.Sheets("Paths").Range("A" & Start).Value
            
            Attached_File = Estimator & "\Spin Notification\" & ControlPanel.Sheets("Paths").Range("E" & Start).Value
            MyMail.Attachments.Add Attached_File
            
            MyMail.Send
        End If
    Next Start
    End If
End Sub
