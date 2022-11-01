Sub termscheck()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim ControlPanel As ThisWorkbook
    Set ControlPanel = ThisWorkbook

    Dim PriceFile As Workbook
    Dim Terms As Workbook
    Dim CopperPriceFile As Workbook
    Dim ProductDatabase As Workbook
    Dim FpeTermsPath As String
    Dim Estimator As String
    Dim AccountNumber As String
    Dim AccountName As String
    Dim PathToPf As String
    Dim PathToCopper As String
    Dim LoopFinish As Integer
    Dim LoopStart As Integer
    Dim Email As String
    Dim MyOutlook As Object
    Dim MyMail As Object
    
    Set MyOutlook = CreateObject("Outlook.Application")
    Set MyMail = MyOutlook.CreateItem(olMailItem)

    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    
    If ControlPanel.Sheets("Terms").AutoFilterMode = True Then
        ControlPanel.Sheets("terms").AutoFilterMode = False
    End If
    
    FpeTermsPath = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\Andrei Polin\Business Objects Terms\Business Objects Terms - FPE.xlsx"
    Set Terms = Workbooks.Open(FpeTermsPath)
    Terms.Sheets(1).Name = "Terms"
    Terms.Sheets("Terms").Rows(1).EntireRow.Delete
    Terms.Sheets("Terms").Rows(1).EntireRow.Delete
    Terms.Sheets("Terms").Columns(1).EntireColumn.Delete
    LastRowTerms = Terms.Sheets("Terms").Range("A" & Rows.Count).End(xlUp).Row
    Terms.Sheets("Terms").Range("A1:S" & LastRowTerms).AutoFilter Field:=14, Criteria1:="<>"
    If Terms.Sheets("Terms").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Terms.Sheets("Terms").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
    End If
    Terms.Sheets("Terms").AutoFilter.ShowAllData
    LastRowTerms = Terms.Sheets("Terms").Range("A" & Rows.Count).End(xlUp).Row
    Terms.Sheets("Terms").Range("A1:S" & LastRowTerms).AutoFilter Field:=13, Criteria1:="="
    If Terms.Sheets("Terms").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Terms.Sheets("Terms").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
    End If
    Terms.Sheets("Terms").AutoFilter.ShowAllData
    LastRowTerms = Terms.Sheets("Terms").Range("A" & Rows.Count).End(xlUp).Row
    Terms.Sheets("Terms").Columns("M:M").TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    LastRowControlPanel = ControlPanel.Sheets("Terms").Range("E" & Rows.Count).End(xlUp).Row
    Terms.Sheets("Terms").Range("T1").Value = "Temporary"
    If Terms.Sheets("Terms").AutoFilterMode = True Then
        Terms.Sheets("Terms").AutoFilterMode = False
    End If
    Terms.Sheets("Terms").Range("A1:T" & LastRowTerms).AutoFilter
    Set ProductDatabase = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Product Database\" & FutureMonthNameDir & "\Product Database.xlsb", ReadOnly:=True)
    ProductDatabase.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    LoopFinish = LastRowControlPanel
    
    For LoopStart = 2 To LoopFinish
        AccountNumber = ControlPanel.Sheets("Terms").Range("A" & LoopStart).Value
        AccountName = ControlPanel.Sheets("Terms").Range("B" & LoopStart).Value
        AccountManager = ControlPanel.Sheets("Terms").Range("C" & LoopStart).Value
        Estimator = ControlPanel.Sheets("Terms").Range("D" & LoopStart).Value
        PathToPf = ControlPanel.Sheets("Terms").Range("E" & LoopStart).Value
        PathToCopper = ControlPanel.Sheets("Terms").Range("F" & LoopStart).Value
        Email = ControlPanel.Sheets("Terms").Range("G" & LoopStart).Value
        
        Set PriceFile = Workbooks.Open(PathToPf, ReadOnly:=True)
        If PathToCopper <> "No Copper" Then
            Set CopperPriceFile = Workbooks.Open(PathToCopper, ReadOnly:=True)
        End If
        Terms.Sheets("Terms").Range("A1:T" & LastRowTerms).AutoFilter Field:=1, Criteria1:=AccountNumber
        If Terms.Sheets("Terms").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            If PathToCopper <> "No Copper" Then
                Terms.Sheets("Terms").Range("T2:T" & LastRowTerms).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(IFERROR(IF(VLOOKUP(@M:M,'[" & PriceFile.Name & "]Price File'!$A:$A,1,FALSE)=@M:M,""Price File"",""Not on Price File""),IF(VLOOKUP(@M:M,'[" & CopperPriceFile.Name & "]Price File'!$A:$A,1,FALSE)=@M:M,""Price File"",""Not on Price File"")),""Not on Price File"")"
            Else
                Terms.Sheets("Terms").Range("T2:T" & LastRowTerms).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(IF(VLOOKUP(@M:M,'[" & PriceFile.Name & "]Price File'!$A:$A,1,FALSE)=@M:M,""Price File"",""Not on Price File""),""Not on Price File"")"
            End If
            Terms.Sheets("Terms").Range("A1:T" & LastRowTerms).AutoFilter Field:=20, Criteria1:="Price File"
            If Terms.Sheets("Terms").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Terms.Sheets("Terms").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            End If
            Terms.Sheets("Terms").AutoFilter.ShowAllData
            LastRowTerms = Terms.Sheets("Terms").Range("A" & Rows.Count).End(xlUp).Row
            Terms.Sheets("Terms").Range("A1:T" & LastRowTerms).AutoFilter Field:=20, Criteria1:="Not on Price File"
            If Terms.Sheets("Terms").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Terms.Sheets.Add.Name = "Temporary"
                Terms.Sheets("Terms").Range("A1:S" & LastRowTerms).Copy
                Terms.Sheets("Temporary").Range("A1").PasteSpecial Paste:=xlPasteValues
                Terms.Sheets("Temporary").Range("A1").PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
                LastRowTemporary = Terms.Sheets("Temporary").Range("A" & Rows.Count).End(xlUp).Row
                Terms.Sheets("Temporary").Range("T1").Value = "LCC Tag"
                Terms.Sheets("Temporary").Range("T2:T" & LastRowTemporary).Formula = "=VLOOKUP(@M:M,'[Product Database.xlsb]Product File (Pyr1)'!$A:$I,9,FALSE)"
                Terms.Sheets("Temporary").Range("T2:T" & LastRowTemporary).Copy
                Terms.Sheets("Temporary").Range("T2:T" & LastRowTemporary).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                Terms.Sheets("Terms").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
                Terms.Sheets("Temporary").Range("A1:T" & LastRowTemporary).AutoFilter
                Terms.Sheets("Temporary").Columns("A:T").AutoFit
                Terms.Sheets("Temporary").Copy
                ActiveWorkbook.SaveAs "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\Andrei Polin\Business Objects Terms\" & FutureMonthNameDir & "\" & AccountNumber & " - " & AccountName & " - Missing Terms.xlsx", FileFormat:=51
                ActiveWorkbook.Close True
                'Set MyMail = MyOutlook.CreateItem(olMailItem)
                'MyMail.To = Email
                'MyMail.Subject = "Important: Missing Terms - " & AccountName & " - " & AccountNumber
                'MyMail.Body = "Hi, " & Estimator & " ,please see attached, Following lines are not on " & AccountName & " Price File. Can you please contact " & AccountManager & " and get Products Removed from terms or added to price file (Re-generate Margin Sheet), including OBS/EOL products, else these products cannot be tracked for SLP Movement."
                'Attached_File = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\Andrei Polin\Business Objects Terms\" & FutureMonthNameDir & "\" & AccountNumber & " - " & AccountName & " - Missing Terms.xlsx"
                'MyMail.Attachments.Add Attached_File
                'MyMail.Send
                Terms.Sheets("Temporary").Delete
            End If
            Terms.Sheets("Terms").AutoFilter.ShowAllData
            If PathToCopper = "No Copper" Then
                PriceFile.Close False
            Else
                PriceFile.Close False
                CopperPriceFile.Close False
            End If
        End If
    Next LoopStart
    ProductDatabase.Close False
    MsgBox "Completed All Accounts"
End Sub
