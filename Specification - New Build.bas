Sub CustomerSpecification()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.Zoom = 85
    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False
    ActiveWindow.DisplayGridlines = False
    
    Dim SpecSheet As ThisWorkbook
    Set SpecSheet = ThisWorkbook
    Dim NewCustomerFile As Workbook
    Dim OldCustomerFile As Workbook
    Dim CurrentSpecification As Workbook
    Dim Estimator As String
    Dim AccountName As String
    Dim AccountNr As String
    
    AccountName = SpecSheet.Sheets("Settings").Range("B2").Value
    AccountNr = SpecSheet.Sheets("Settings").Range("B3").Value
    Estimator = SpecSheet.Sheets("Settings").Range("B1").Value
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    CustomerSpecificationPath = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Templates\" & AccountNr & " Quote.xlsx"
    ActiveYear = Format(DateSerial(Year(Date), Month(Date), 1), "yyyy")
    FutureYear = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "yyyy")
    
    Set NewCustomerFile = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Customer Price Files\" & FutureMonthNameDir & "\" & AccountName & " - " & AccountNr & " Customer Version.xlsx")
    BLOCK = SpecSheet.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
    LastRowSpecSheet = SpecSheet.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
    LastRowNewCustomerFile = NewCustomerFile.Sheets("Customer Version").Range("A" & Rows.Count).End(xlUp).Row
    
    If LastRowSpecSheet = 1 Then
        NewCustomerFile.Sheets("Customer Version").Range("A12:B" & LastRowNewCustomerFile).Copy
        SpecSheet.Sheets("Main").Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        NewCustomerFile.Sheets("Customer Version").Range("E12:E" & LastRowNewCustomerFile).Copy
        SpecSheet.Sheets("Main").Range("C2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        NewCustomerFile.Sheets("Customer Version").Range("D12:D" & LastRowNewCustomerFile).Copy
        SpecSheet.Sheets("Main").Range("D2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowSpecSheet = SpecSheet.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        SpecSheet.Sheets("Main").Range("E2:E" & LastRowSpecSheet).Value = "Unknown"
        SpecSheet.Sheets("Main").Range("F2:F" & LastRowSpecSheet).Value = "0"
        SpecSheet.Sheets("Main").Range("G2:G" & LastRowSpecSheet).Value = "TRUE"
    Else
        LastRowSpecSheet = SpecSheet.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
        SpecSheet.Sheets("Main").Range("H2:H" & LastRowSpecSheet).Formula = "=VLOOKUP(@A:A,'[" & NewCustomerFile.Name & "]Customer Version'!$A:$A,1,FALSE)"
        SpecSheet.Sheets("Main").Range("A1:H" & LastRowSpecSheet).AutoFilter field:=8, Criteria1:="#N/A"
            If SpecSheet.Sheets("Main").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                SpecSheet.Sheets("Main").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
                SpecSheet.Sheets("Main").AutoFilter.ShowAllData
                SpecSheet.Sheets("Main").Range("H2:H" & LastRowSpecSheet).ClearContents
            Else
                SpecSheet.Sheets("Main").AutoFilter.ShowAllData
                SpecSheet.Sheets("Main").Range("H2:H" & LastRowSpecSheet).ClearContents
            End If
        LastRowSpecSheet = SpecSheet.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
        NewCustomerFile.Sheets("Customer Version").Range("AE12:AE" & LastRowNewCustomerFile).Formula = "=VLOOKUP(@A:A,'[" & SpecSheet.Name & "]Main'!$A:$A,1,FALSE)"
        NewCustomerFile.Sheets("Customer Version").AutoFilterMode = False
        NewCustomerFile.Sheets("Customer Version").Range("A11:AE" & LastRowNewCustomerFile).AutoFilter field:=31, Criteria1:="#N/A"
            If NewCustomerFile.Sheets("Customer Version").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                NewCustomerFile.Sheets("Customer Version").Range("A12:B" & LastRowNewCustomerFile).SpecialCells(xlCellTypeVisible).Copy
                SpecSheet.Sheets("Main").Range("A" & LastRowSpecSheet + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                NewCustomerFile.Sheets("Customer Version").Range("E12:E" & LastRowNewCustomerFile).SpecialCells(xlCellTypeVisible).Copy
                SpecSheet.Sheets("Main").Range("C" & LastRowSpecSheet + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                NewCustomerFile.Sheets("Customer Version").Range("D12:D" & LastRowNewCustomerFile).SpecialCells(xlCellTypeVisible).Copy
                SpecSheet.Sheets("Main").Range("D" & LastRowSpecSheet + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                LastRowSpecSheetNew = SpecSheet.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
                SpecSheet.Sheets("Main").Range("E" & LastRowSpecSheet + 1 & ":E" & LastRowSpecSheetNew).Value = "Unknown"
                SpecSheet.Sheets("Main").Range("F" & LastRowSpecSheet + 1 & ":F" & LastRowSpecSheetNew).Value = "0"
                SpecSheet.Sheets("Main").Range("G" & LastRowSpecSheet + 1 & ":G" & LastRowSpecSheetNew).Value = "TRUE"
            End If
        NewCustomerFile.Sheets("Customer Version").AutoFilterMode = False
        LastRowSpecSheet = SpecSheet.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
    End If
    
    If LastRowSpecSheet > BLOCK Then
        MsgBox "Detected new products on " & AccountName & ", Please check SpecSheet for changes and re-run the script, script will now exit."
        NewCustomerFile.Close False
    Else
        SpecSheet.Save
        SpecSheet.Sheets("Main").Range("A1:G" & LastRowSpecSheet).AutoFilter field:=7, Criteria1:="FALSE"
        If SpecSheet.Sheets("Main").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            SpecSheet.Sheets("main").AutoFilterMode = False
            Set CurrentSpecification = Workbooks.Open(CustomerSpecificationPath)
            SpecSheet.Sheets("Main").Range("A1:G" & LastRowSpecSheet).Copy
            CurrentSpecification.Sheets("Temp").Range("A1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowCurrentSpecificationTemp = CurrentSpecification.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
                If ActiveMonthName = "January" Then
                        NewCustomerFile.Sheets("Customer Version").Range("J:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "February" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:G,L:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "March" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:I,N:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "April" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:K,P:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "May" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:M,R:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "June" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:O,T:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "July" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:Q,V:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "August" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:S,X:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "September" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:U,Z:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "October" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:W,AB:AC").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "November" Then
                        NewCustomerFile.Sheets("Customer Version").Range("F:Y").EntireColumn.Delete
                    Else
                End If
                If ActiveMonthName = "December" Then
                        NewCustomerFile.Sheets("Customer Version").Range("H:AA").EntireColumn.Delete
                    Else
                End If
            CurrentSpecification.Sheets("Temp").Range("H1:H" & LastRowCurrentSpecificationTemp).Formula = "=VLOOKUP(@A:A,'[" & NewCustomerFile.Name & "]Customer Version'!$A:$AX,6,FALSE)"
            CurrentSpecification.Sheets("Temp").Range("I1:I" & LastRowCurrentSpecificationTemp).Formula = "=VLOOKUP(@A:A,'[" & NewCustomerFile.Name & "]Customer Version'!$A:$AX,8,FALSE)"
            CurrentSpecification.Sheets("Temp").Range("H1:I" & LastRowCurrentSpecificationTemp).Copy
            CurrentSpecification.Sheets("Temp").Range("H1:I" & LastRowCurrentSpecificationTemp).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            NewCustomerFile.Close False
            CurrentSpecification.Sheets("Temp").Range("A1:I" & LastRowCurrentSpecificationTemp).AutoFilter field:=7, Criteria1:=True
            If CurrentSpecification.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                CurrentSpecification.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
                CurrentSpecification.Sheets("Temp").AutoFilter.ShowAllData
                LastRowCurrentSpecificationTemp = CurrentSpecification.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
            Else
                CurrentSpecification.Sheets("Temp").AutoFilter.ShowAllData
            End If
            
            If LastRowCurrentSpecificationTemp > 0 Then
                CurrentSpecification.Sheets("Temp").Range("E2:F" & LastRowCurrentSpecificationTemp).Copy
                CurrentSpecification.Sheets("Clean").Range("A1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                CurrentSpecification.Sheets("Clean").Columns("A:B").RemoveDuplicates Columns:=1, Header:=xlNo
                LastRowCurrentSpecificationC = CurrentSpecification.Sheets("Clean").Range("A" & Rows.Count).End(xlUp).Row
                CurrentSpecification.Sheets("Clean").Range("A1:B" & LastRowCurrentSpecificationC).Sort Key1:=CurrentSpecification.Sheets("Clean").Range("B1:B" & LastRowCurrentSpecificationC), Order1:=xlAscending, Header:=xlNo
                CurrentSpecification.Sheets("Quote").Range("F8").EntireRow.Resize(LastRowCurrentSpecificationC).Insert Shift:=xlDown
                CurrentSpecification.Sheets("Clean").Range("A1:A" & LastRowCurrentSpecificationC).Copy
                CurrentSpecification.Sheets("Quote").Range("F8").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                LastRowCurrentSpecificationTemp = CurrentSpecification.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
                LastRowCurrentSpecificationMain = CurrentSpecification.Sheets("Quote").Range("A" & Rows.Count).End(xlUp).Row + 2
                FirstRow = CurrentSpecification.Sheets("Quote").Range("A" & Rows.Count).End(xlUp).Row + 2
                SubheadingStart = 1
                Subheadingend = LastRowCurrentSpecificationC
                Value = 1
                
                For I = SubheadingStart To Subheadingend
                    SearchHeading = CurrentSpecification.Sheets("Clean").Range("A" & Value).Value
                    CurrentSpecification.Sheets("Temp").Range("A1:G" & LastRowCurrentSpecificationTemp).AutoFilter field:=5, Criteria1:=SearchHeading
                        If CurrentSpecification.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                            LastRowCurrentSpecificationMain = CurrentSpecification.Sheets("Quote").Range("A" & Rows.Count).End(xlUp).Row + 2
                            CurrentSpecification.Sheets("Quote").Range("F" & LastRowCurrentSpecificationMain).Value = SearchHeading
                            CurrentSpecification.Sheets("Temp").Range("A2:A" & LastRowCurrentSpecificationTemp).SpecialCells(xlCellTypeVisible).Copy
                            CurrentSpecification.Sheets("Quote").Range("A" & LastRowCurrentSpecificationMain + 1).PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                            CurrentSpecification.Sheets("Temp").Range("D2:D" & LastRowCurrentSpecificationTemp).SpecialCells(xlCellTypeVisible).Copy
                            CurrentSpecification.Sheets("Quote").Range("B" & LastRowCurrentSpecificationMain + 1).PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                            CurrentSpecification.Sheets("Temp").Range("B2:B" & LastRowCurrentSpecificationTemp).SpecialCells(xlCellTypeVisible).Copy
                            CurrentSpecification.Sheets("Quote").Range("C" & LastRowCurrentSpecificationMain + 1).PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                            CurrentSpecification.Sheets("Temp").Range("H2:H" & LastRowCurrentSpecificationTemp).SpecialCells(xlCellTypeVisible).Copy
                            CurrentSpecification.Sheets("Quote").Range("I" & LastRowCurrentSpecificationMain + 1).PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                            CurrentSpecification.Sheets("Temp").Range("I2:I" & LastRowCurrentSpecificationTemp).SpecialCells(xlCellTypeVisible).Copy
                            CurrentSpecification.Sheets("Quote").Range("J" & LastRowCurrentSpecificationMain + 1).PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                            LastRowCurrentSpecificationMain = CurrentSpecification.Sheets("Quote").Range("A" & Rows.Count).End(xlUp).Row + 2
                        End If
                    Value = Value + 1
                Next I
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).AutoFilter field:=6, Criteria1:="<>"
                    If CurrentSpecification.Sheets("Quote").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(205, 46, 46)
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Font.Color = vbWhite
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Font.Bold = True
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlEdgeTop).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlEdgeRight).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).HorizontalAlignment = xlCenter
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).VerticalAlignment = xlCenter
                        CurrentSpecification.Sheets("Quote").AutoFilter.ShowAllData
                    End If
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).AutoFilter field:=1, Criteria1:="<>"
                    If CurrentSpecification.Sheets("Quote").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":A" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).HorizontalAlignment = xlLeft
                        CurrentSpecification.Sheets("Quote").AutoFilter.ShowAllData
                    End If
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).AutoFilter field:=6, Criteria1:="="
                    If CurrentSpecification.Sheets("Quote").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlEdgeRight).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":C" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlInsideVertical).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").Range("H" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Borders(xlInsideVertical).LineStyle = xlContinuous
                        CurrentSpecification.Sheets("Quote").AutoFilter.ShowAllData
                    End If
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).AutoFilter field:=1, Criteria1:="<>"
                    If CurrentSpecification.Sheets("Quote").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                        Set OldCustomerFile = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Customer Price Files\" & ActiveMonthNameDir & "\" & AccountName & " - " & AccountNr & " Customer Version.xlsx")
                        LastRowOldCustomerFile = OldCustomerFile.Sheets("Customer Version").Range("A" & Rows.Count).End(xlUp).Row
                        CurrentSpecification.Sheets("Quote").Range("K" & FirstRow + 1 & ":K" & LastRowCurrentSpecificationMain - 2).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(IF(@A:A=VLOOKUP(@A:A,'[" & OldCustomerFile.Name & "]Customer Version'!$A:$A,1,FALSE),"""",""Addition""),""Addition"")"
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow & ":K" & LastRowCurrentSpecificationMain).AutoFilter field:=11, Criteria1:="<>Addition"
                        If CurrentSpecification.Sheets("Quote").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                            CurrentSpecification.Sheets("Quote").Range("K" & FirstRow + 1 & ":K" & LastRowCurrentSpecificationMain - 2).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(IF(@J:J>@I:I,""Increase"",IF(@J:J<@I:I,""Decrease"","""")),""ERROR"")"
                            CurrentSpecification.Sheets("Quote").AutoFilter.ShowAllData
                        End If
                        CurrentSpecification.Sheets("Quote").AutoFilter.ShowAllData
                        CurrentSpecification.Sheets("Quote").Range("I" & FirstRow & ":J" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Style = "Currency"
                        CurrentSpecification.Sheets("Quote").Range("K" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).Copy
                        CurrentSpecification.Sheets("Quote").Range("K" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                        CurrentSpecification.Sheets("Quote").Range("I" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).HorizontalAlignment = xlCenter
                        CurrentSpecification.Sheets("Quote").Range("I" & FirstRow & ":K" & LastRowCurrentSpecificationMain).SpecialCells(xlCellTypeVisible).VerticalAlignment = xlCenter
                        CurrentSpecification.Sheets("Quote").Range("A" & FirstRow - 2).Formula = "=COUNTIF(K:K,""Increase"") & "" Increase(s)"""
                        CurrentSpecification.Sheets("Quote").Range("C" & FirstRow - 2).Formula = "=COUNTIF(K:K,""Decrease"") & "" Decrease(s)"""
                        CurrentSpecification.Sheets("Quote").Range("G" & FirstRow - 2).Formula = "=COUNTIF(K:K,""Addition"") & "" Addition(s)"""
                        CurrentSpecification.Sheets("Quote").AutoFilter.ShowAllData
                        OldCustomerFile.Sheets("Customer Version").Range("AE12:AE" & LastRowOldCustomerFile).Formula = "=VLOOKUP(@A:A,'[" & CurrentSpecification.Name & "]Temp'!$A:$A,1,FALSE)"
                        OldCustomerFile.Sheets("Customer Version").Range("A11:AE" & LastRowOldCustomerFile).AutoFilter
                        OldCustomerFile.Sheets("Customer Version").Range("A11:AE" & LastRowOldCustomerFile).AutoFilter field:=31, Criteria1:="#N/A"
                        If OldCustomerFile.Sheets("Customer Version").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                           OldCustomerFile.Sheets("Customer Version").Range("A12:E" & LastRowOldCustomerFile).SpecialCells(xlCellTypeVisible).Copy
                           CurrentSpecification.Sheets("Withdrawn").Range("A1").PasteSpecial Paste:=xlPasteValues
                           CurrentSpecification.Sheets("Quote").Range("J" & FirstRow - 2).Formula = "=COUNTA(Withdrawn!A:A) & "" Withdrawn(s)"""
                        Else
                            CurrentSpecification.Sheets("Quote").Range("J" & FirstRow - 2).Formula = "=COUNTA(Withdrawn!A:A) & "" Withdrawn(s)"""
                        End If
                    End If
                OldCustomerFile.Close False
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow - 2 & ":K" & FirstRow - 2).Copy
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow - 2 & ":K" & FirstRow).Copy
                CurrentSpecification.Sheets("Quote").Range("A" & FirstRow - 2 & ":K" & FirstRow).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                CurrentSpecification.Sheets("Quote").Range("F4").Formula = "=I6"
                CurrentSpecification.Sheets("Quote").Range("I6").Value = ActiveMonthName & " " & ActiveYear
                CurrentSpecification.Sheets("Quote").Range("J6").Value = FutureMonthName & " " & FutureYear
                CurrentSpecification.Sheets("Quote").Range("C4").Formula = "=J6"
                CurrentSpecification.Sheets("Quote").Activate
                CurrentSpecification.Sheets("Quote").Range("F8").End(xlDown).Select
                MenuLastRow = ActiveCell.Row
                SubheadingStart = 8
                Subheadingend = MenuLastRow
                Value = 8
                
                CurrentSpecification.Sheets("Quote").AutoFilterMode = False
                For I = SubheadingStart To MenuLastRow
                    SearchHeading = CurrentSpecification.Sheets("Quote").Range("F" & Value).Value
                    CurrentSpecification.Sheets("Quote").Range("A" & (FirstRow - 3) & ":K" & LastRowCurrentSpecificationMain - 2).AutoFilter field:=6, Criteria1:="*" & SearchHeading & "*"
                        If CurrentSpecification.Sheets("Quote").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                            CurrentSpecification.Activate
                            CurrentSpecification.Sheets("Quote").Range("F" & FirstRow - 3).Select
                            Range("F" & (FirstRow - 3) & ":F" & LastRowCurrentSpecificationMain).Offset(1, 0).SpecialCells(xlCellTypeVisible).Areas(1).Rows(1).Select
                            LocationS = ActiveCell.Address
                            CurrentSpecification.Sheets("Quote").Range("F" & Value).Select
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LocationS, TextToDisplay:=SearchHeading
                        End If
                    Value = Value + 1
                Next I
                
                Set SPGFile = Workbooks.Open(SpecSheet.Path & "\" & AccountNr & " SPG.xlsx")
                LastRowSPGFile = SPGFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
                CurrentSpecification.Sheets("Quote").AutoFilterMode = False
                CurrentSpecification.Sheets("Quote").Range("F" & LastRowCurrentSpecificationMain).EntireRow.Resize(LastRowSPGFile + 3).Insert Shift:=xlDown
                CurrentSpecification.Sheets("Quote").Range("C" & LastRowCurrentSpecificationMain & ":H" & LastRowCurrentSpecificationMain).Merge
                CurrentSpecification.Sheets("Quote").Range("C" & LastRowCurrentSpecificationMain & ":H" & LastRowCurrentSpecificationMain).Value = "Please note that these plastics discounts are to be used as a guideline only.  This is the minimum discount the contractor will be expected to purchase at.  Support is on a site by site basis so please ensure that you set this up at the branch."
                CurrentSpecification.Sheets("Quote").Rows(LastRowCurrentSpecificationMain & ":" & LastRowCurrentSpecificationMain).RowHeight = 69
                CurrentSpecification.Sheets("Quote").Range("C" & LastRowCurrentSpecificationMain & ":H" & LastRowCurrentSpecificationMain).WrapText = True
                CurrentSpecification.Sheets("Quote").Range("C" & LastRowCurrentSpecificationMain & ":H" & LastRowCurrentSpecificationMain).Interior.Color = RGB(255, 255, 0)
                CurrentSpecification.Sheets("Quote").Range("C" & LastRowCurrentSpecificationMain & ":H" & LastRowCurrentSpecificationMain).Font.Color = vbBlack
                
                SPGFile.Sheets(1).Range("A1:A" & LastRowSPGFile).Copy
                CurrentSpecification.Sheets("Quote").Range("B" & LastRowCurrentSpecificationMain + 2).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                SPGFile.Sheets(1).Range("B1:B" & LastRowSPGFile).Copy
                CurrentSpecification.Sheets("Quote").Range("F" & LastRowCurrentSpecificationMain + 2).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                SPGFile.Sheets(1).Range("C1:C" & LastRowSPGFile).Copy
                CurrentSpecification.Sheets("Quote").Range("I" & LastRowCurrentSpecificationMain + 2).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                SPGFile.Close False
                CurrentSpecification.Sheets("Quote").AutoFilterMode = False
                CurrentSpecification.Sheets("Quote").Columns("I:J").AutoFit
                CurrentSpecification.Sheets("Quote").Range("I" & LastRowCurrentSpecificationMain & ":I" & LastRowCurrentSpecificationMain + LastRowSPGFile + 2).NumberFormat = "0.00%"
                CurrentSpecification.Activate
                CurrentSpecification.Sheets("Quote").Range("A1").Select
                CurrentSpecification.Sheets("Quote").Rows(7).Select
                ActiveWindow.FreezePanes = True
                CurrentSpecification.Sheets("Temp").Delete
                CurrentSpecification.Sheets("Clean").Delete
                LastRowCurrentSpecification = CurrentSpecification.Sheets("Quote").Range("A" & Rows.Count).End(xlUp).Row
                CurrentSpecification.Sheets("Quote").PageSetup.PrintArea = CurrentSpecification.Sheets("Quote").Range("A1:K" & LastRowCurrentSpecification + 3)
                
                
                Dim PathDir As String
                Dim PathDir2 As String
                Dim FileNameSpec As String
                PathDir = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\"
                PathDir2 = "\Customer Specification\"
                FileNameSpec = AccountNr & " Quote (" & FutureMonthName & " " & FutureYear & ").xls"
                Application.DisplayAlerts = False
                CurrentSpecification.SaveAs Filename:=PathDir & Estimator & PathDir2 & FutureMonthNameDir & "\" & FileNameSpec, FileFormat:=56
                CurrentSpecification.Close False
            End If
        Else
            MsgBox AccountName & " - There are no products to add on Specification file, please check SpecSheet, Script will now exit"
            SpecSheet.Sheets("Main").AutoFilterMode = False
            NewCustomerFile.Close False
        End If
    End If
    SpecSheet.Save
End Sub