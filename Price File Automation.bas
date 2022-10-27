Sub Price_File_Automation_PrePP()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook
    
    Dim Additions As Workbook
    Dim PricePointUpload As Workbook
    Dim ProductDatabase As Workbook
    Dim ManualSupports As Workbook
    Dim StrFileName As String
    Dim StrFileExists As String
    
    Master.Sheets("Price File").Columns.EntireColumn.Hidden = False
    Master.Sheets("Price File").Rows.EntireRow.Hidden = False
    
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    CurrentAccount = Master.Sheets("Price File").Range("B3").Value
    Estimator = Master.Sheets("Settings").Range("B1").Value
    Pricefiletype = Master.Sheets("Settings").Range("B7").Value
    
    If Master.Sheets("Price File").AutoFilterMode = True Then
        Master.Sheets("Price File").AutoFilterMode = False
    End If
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    Master.Sheets("Price File").Range("A11:AQ" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("A11:AQ" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    If LastRowMaster > 11 Then
        answer = MsgBox("Do you want to clear column Manual Margin Field?", vbQuestion + vbYesNo + vbDefaultButton2, "Attention")
        If answer = vbYes Then
            If Pricefiletype = "MARGIN" Then
            Master.Sheets("Price File").Range("AV12:AW" & LastRowMaster).ClearContents
            Master.Sheets("Price File").Range("BW12:BX" & LastRowMaster).ClearContents
            If Master.Sheets("Settings").Range("B9").Value <> "YES" Then
                Master.Sheets("Price File").Range("AV12:AV" & LastRowMaster).Value = "X"
            End If
            Master.Sheets("Price File").Range("BK12:BM" & LastRowMaster).ClearContents
        ElseIf Pricefiletype = "SLP" Then
            Master.Sheets("Price File").Range("AV12:AW" & LastRowMaster).ClearContents
            Master.Sheets("Price File").Range("BW12:BX" & LastRowMaster).ClearContents
            Master.Sheets("Price File").Range("BK12:BM" & LastRowMaster).ClearContents
            End If
        End If
    End If
    
    Set Additions = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Price File Additions.xlsx")
    If Additions.Sheets(1).AutoFilterMode = True Then
        Additions.Sheets(1).AutoFilterMode = False
    End If
    Additions.Sheets(1).Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("Price File").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    LastRowAdditions = Additions.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Additions.Sheets(1).Range("A1:C" & LastRowAdditions).AutoFilter Field:=1, Criteria1:=CurrentAccount
    If Additions.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Additions.Sheets.Add(After:=Sheets(1)).Name = "Temp"
        Additions.Sheets(1).Range("A1:C" & LastRowAdditions).Copy
        Additions.Sheets("Temp").Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").Range("D1").Value = "Temp"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Formula = "=IF(COUNTIF(B:B,B2)>1,IF(COUNTIF(B$2:B2,B2)=1,""x"",""xx""),"""")"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Copy
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Additions.Sheets("Temp").Range("A1:D" & LastRowTempo).AutoFilter Field:=4, Criteria1:="x"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilter.ShowAllData
        End If
        Additions.Sheets("Temp").AutoFilter.ShowAllData
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Formula = "=IF(COUNTIF(B:B,B2)>1,IF(COUNTIF(B$2:B2,B2)=1,""x"",""xx""),"""")"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Copy
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Additions.Sheets("Temp").Range("A1:D" & LastRowTempo).AutoFilter Field:=4, Criteria1:="x"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilter.ShowAllData
        End If
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Formula = "=IF(COUNTIF(B:B,B2)>1,IF(COUNTIF(B$2:B2,B2)=1,""x"",""xx""),"""")"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Copy
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Additions.Sheets("Temp").Range("A1:D" & LastRowTempo).AutoFilter Field:=4, Criteria1:="x"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilter.ShowAllData
        End If
        Additions.Sheets("Temp").AutoFilter.ShowAllData
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Formula = "=IF(COUNTIF(B:B,B2)>1,IF(COUNTIF(B$2:B2,B2)=1,""x"",""xx""),"""")"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Copy
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Additions.Sheets("Temp").Range("A1:D" & LastRowTempo).AutoFilter Field:=4, Criteria1:="x"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilter.ShowAllData
        End If
        Additions.Sheets("Temp").AutoFilter.ShowAllData
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Formula = "=IF(COUNTIF(B:B,B2)>1,IF(COUNTIF(B$2:B2,B2)=1,""x"",""xx""),"""")"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Copy
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Additions.Sheets("Temp").Range("A1:D" & LastRowTempo).AutoFilter Field:=4, Criteria1:="x"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilter.ShowAllData
        End If
        Additions.Sheets("Temp").AutoFilter.ShowAllData
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Formula = "=IF(COUNTIF(B:B,B2)>1,IF(COUNTIF(B$2:B2,B2)=1,""x"",""xx""),"""")"
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).Copy
        Additions.Sheets("Temp").Range("D2:D" & LastRowTempo).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Additions.Sheets("Temp").Range("A1:D" & LastRowTempo).AutoFilter Field:=4, Criteria1:="x"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilter.ShowAllData
        End If
        Additions.Sheets("Temp").AutoFilter.ShowAllData
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").Range("D1:D" & LastRowTempo).ClearContents
        LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        Additions.Sheets("Temp").AutoFilterMode = False
        
        If ActiveMonthName = "January" Then
            SELL_PRICE_CURRENT = "G12:G"
            SELL_PRICE_FUTURE = "J12:J"
        End If
        If ActiveMonthName = "February" Then
            SELL_PRICE_CURRENT = "J12:J"
            SELL_PRICE_FUTURE = "M12:M"
        End If
        If ActiveMonthName = "March" Then
            SELL_PRICE_CURRENT = "M12:M"
            SELL_PRICE_FUTURE = "P12:P"
        End If
        If ActiveMonthName = "April" Then
            SELL_PRICE_CURRENT = "P12:P"
            SELL_PRICE_FUTURE = "S12:S"
        End If
        If ActiveMonthName = "May" Then
            SELL_PRICE_CURRENT = "S12:S"
            SELL_PRICE_FUTURE = "V12:V"
        End If
        If ActiveMonthName = "June" Then
            SELL_PRICE_CURRENT = "V12:V"
            SELL_PRICE_FUTURE = "Y12:Y"
        End If
        If ActiveMonthName = "July" Then
            SELL_PRICE_CURRENT = "Y12:Y"
            SELL_PRICE_FUTURE = "AB12:AB"
        End If
        If ActiveMonthName = "August" Then
            SELL_PRICE_CURRENT = "AB12:AB"
            SELL_PRICE_FUTURE = "AE12:AE"
        End If
        If ActiveMonthName = "September" Then
            SELL_PRICE_CURRENT = "AE12:AE"
            SELL_PRICE_FUTURE = "AH12:AH"
        End If
        If ActiveMonthName = "October" Then
            SELL_PRICE_CURRENT = "AH12:AH"
            SELL_PRICE_FUTURE = "AK12:AK"
        End If
        If ActiveMonthName = "November" Then
            SELL_PRICE_CURRENT = "AK12:AK"
            SELL_PRICE_FUTURE = "AN12:AN"
        End If
        If ActiveMonthName = "December" Then
            SELL_PRICE_CURRENT = "AN12:AN"
            SELL_PRICE_FUTURE = "G12:G"
        End If
        
        If LastRowMaster > 11 Then
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).Formula = "=IFERROR(TRUNC(VLOOKUP(@A:A,'[Price File Additions.xlsx]Temp'!$B:$C,2,FALSE),2),0)"
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).Copy
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Price File").Range("A11:BZ" & LastRowMaster).AutoFilter Field:=77, Criteria1:=">0"
            If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Master.Sheets("Price File").Range(SELL_PRICE_CURRENT & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=@BY:BY"
                Master.Sheets("Price File").Range(SELL_PRICE_FUTURE & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=@BY:BY"
                Master.Sheets("Price File").AutoFilter.ShowAllData
                Master.Sheets("Price File").Range(SELL_PRICE_CURRENT & LastRowMaster).Copy
                Master.Sheets("Price File").Range(SELL_PRICE_CURRENT & LastRowMaster).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                Master.Sheets("Price File").Range(SELL_PRICE_FUTURE & LastRowMaster).Copy
                Master.Sheets("Price File").Range(SELL_PRICE_FUTURE & LastRowMaster).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            End If
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Additions.Sheets("Temp").Range("A2:A" & LastRowTempo).Formula = "=VLOOKUP(B2,'[" & Master.Name & "]Price File'!$A:$A,1,FALSE)"
        Additions.Sheets("Temp").Range("A1:C" & LastRowTempo).AutoFilter Field:=1, Criteria1:="<>#N/A"
        If Additions.Sheets("Temp").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Additions.Sheets("Temp").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Additions.Sheets("Temp").AutoFilterMode = False
            LastRowTempo = Additions.Sheets("Temp").Range("A" & Rows.Count).End(xlUp).Row
        End If
        If LastRowTempo > 1 Then
            Additions.Sheets("Temp").AutoFilterMode = False
            Additions.Sheets("Temp").Range("B2:B" & LastRowTempo).Copy
            Master.Sheets("Price File").Range("A" & LastRowMaster + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).Formula = "=IFERROR(TRUNC(VLOOKUP(@A:A,'[Price File Additions.xlsx]Temp'!$B:$C,2,FALSE),2),0)"
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).Copy
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Price File").Range("A11:BZ" & LastRowMaster).AutoFilter Field:=77, Criteria1:=">0"
            If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Master.Sheets("Price File").Range(SELL_PRICE_CURRENT & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=@BY:BY"
                Master.Sheets("Price File").Range(SELL_PRICE_FUTURE & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=@BY:BY"
                Master.Sheets("Price File").AutoFilter.ShowAllData
                Master.Sheets("Price File").Range(SELL_PRICE_CURRENT & LastRowMaster).Copy
                Master.Sheets("Price File").Range(SELL_PRICE_CURRENT & LastRowMaster).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                Master.Sheets("Price File").Range(SELL_PRICE_FUTURE & LastRowMaster).Copy
                Master.Sheets("Price File").Range(SELL_PRICE_FUTURE & LastRowMaster).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            End If
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").Range("BY11:BY" & LastRowMaster).ClearContents
        Additions.Worksheets("Temp").Delete
        Additions.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
    End If
    Additions.Sheets(1).AutoFilterMode = False
    Master.Sheets("Price File").AutoFilterMode = False
    If Pricefiletype = "MARGIN" Then
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=2, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("AV12:AV" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "X"
        End If
        Master.Sheets("Price File").AutoFilterMode = False
    End If
    Additions.Close True
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="="
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets("Price File").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
        Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilterMode = False
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    
    If LastRowMaster > 11 Then
        StrFileName = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\PricePoint Upload\" & CurrentAccount & "UPLOAD.xlsx"
        StrFileExists = Dir(StrFileName)
        
        If StrFileExists = "" Then
            Workbooks.Add
            ActiveWorkbook.Sheets(1).Range("A1").Value = "Section"
            ActiveWorkbook.Sheets(1).Range("B1").Value = "Sub-Section"
            ActiveWorkbook.Sheets(1).Range("C1").Value = "Supplier Product Code"
            ActiveWorkbook.Sheets(1).Range("D1").Value = "WUK Product Code"
            ActiveWorkbook.Sheets(1).Range("E1").Value = "Product Comment"
            ActiveWorkbook.Sheets(1).Range("F1").Value = "Quantity"
            ActiveWorkbook.Sheets(1).Range("G1").Value = "Selling Price Ã‚Â£"
            ActiveWorkbook.Sheets(1).Range("H1").Value = "Discount 1 %"
            ActiveWorkbook.Sheets(1).Range("I1").Value = "Discount 2 %"
            ActiveWorkbook.Sheets(1).Range("J1").Value = "Trading Margin %"
            ActiveWorkbook.Sheets(1).Range("K1").Value = "Product Specific Customer Rebate %"
            ActiveWorkbook.Sheets(1).Range("L1").Value = "Supplier Contract"
            ActiveWorkbook.Sheets(1).Range("M1").Value = "Agreed Contract Support %"
            ActiveWorkbook.Sheets(1).Range("N1").Value = "Agreed Contract Support Ã‚Â£"
            ActiveWorkbook.Sheets(1).Range("O1").Value = "Agreed Net Price Ã‚Â£"
            ActiveWorkbook.Sheets(1).Range("D2").Value = "1111111"
            ActiveWorkbook.Sheets(1).Range("D3").Value = "2222222"
            ActiveWorkbook.Sheets(1).Range("D4").Value = "3333333"
            ActiveWorkbook.SaveAs "" & StrFileName & ""
            ActiveWorkbook.Close True
        End If
        
        Set PricePointUpload = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\Pricepoint Upload\" & CurrentAccount & "UPLOAD.xlsx")
        Set ManualSupports = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Manual Supports.xlsx")
        If ManualSupports.Sheets(1).AutoFilterMode = True Then
            ManualSupports.Sheets(1).AutoFilterMode = False
        End If
        If PricePointUpload.Sheets(1).AutoFilterMode = True Then
            PricePointUpload.Sheets(1).AutoFilterMode = False
        End If
        LastRowPricePointUpload = PricePointUpload.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row
        LastRowManualSupports = ManualSupports.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        PricePointUpload.Sheets(1).Range("A2:O" & LastRowPricePointUpload).ClearContents
        Master.Sheets("Price File").Range("A12:A" & LastRowMaster).Copy
        PricePointUpload.Sheets(1).Range("D2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowPricePointUpload = PricePointUpload.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row
        
        ManualSupports.Sheets("Supports").Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        ManualSupports.Sheets("Supports").Range("F1:F" & LastRowManualSupports).NumberFormat = "dd/mm/yyyy"
        ManualSupports.Sheets("Supports").Range("A1:F" & LastRowManualSupports).AutoFilter Field:=6, Criteria1:="<" & CDbl(Date)
        If ManualSupports.Sheets("Supports").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            ManualSupports.Sheets("Supports").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            ManualSupports.Sheets("Supports").AutoFilter.ShowAllData
            LastRowManualSupports = ManualSupports.Sheets("Supports").Range("A" & Rows.Count).End(xlUp).Row
        End If
        ManualSupports.Sheets(1).AutoFilter.ShowAllData
        ManualSupports.Sheets("Supports").Range("A1:D" & LastRowManualSupports).AutoFilter Field:=1, Criteria1:=CurrentAccount
        If ManualSupports.Sheets("Supports").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            ManualSupports.Sheets.Add(After:=Sheets("Supports")).Name = "Temporary"
            ManualSupports.Sheets("Supports").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Copy
            ManualSupports.Sheets("Temporary").Range("A1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowUploadSheet = PricePointUpload.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row
            PricePointUpload.Sheets(1).Range("M2:M" & LastRowUploadSheet).Formula = "=VLOOKUP(@D:D,'[Manual Supports.xlsx]Temporary'!$B:$F,2,FALSE)"
            PricePointUpload.Sheets(1).Range("N2:N" & LastRowUploadSheet).Formula = "=VLOOKUP(@D:D,'[Manual Supports.xlsx]Temporary'!$B:$F,3,FALSE)"
            PricePointUpload.Sheets(1).Range("O2:O" & LastRowUploadSheet).Formula = "=VLOOKUP(@D:D,'[Manual Supports.xlsx]Temporary'!$B:$F,4,FALSE)"
            PricePointUpload.Sheets(1).Range("M2:O" & LastRowUploadSheet).Copy
            PricePointUpload.Sheets(1).Range("M2:O" & LastRowUploadSheet).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            PricePointUpload.Sheets(1).Range("A1:O" & LastRowUploadSheet).AutoFilter Field:=13, Criteria1:="#N/A", Operator:=xlOr, Criteria2:="0"
            If PricePointUpload.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                PricePointUpload.Sheets(1).Range("M2:M" & LastRowUploadSheet).SpecialCells(xlCellTypeVisible).ClearContents
                PricePointUpload.Sheets(1).AutoFilter.ShowAllData
            End If
            PricePointUpload.Sheets(1).AutoFilter.ShowAllData
            PricePointUpload.Sheets(1).Range("A1:O" & LastRowUploadSheet).AutoFilter Field:=14, Criteria1:="#N/A", Operator:=xlOr, Criteria2:="0"
            If PricePointUpload.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                PricePointUpload.Sheets(1).Range("N2:N" & LastRowUploadSheet).SpecialCells(xlCellTypeVisible).ClearContents
                PricePointUpload.Sheets(1).AutoFilter.ShowAllData
            End If
            PricePointUpload.Sheets(1).AutoFilter.ShowAllData
            PricePointUpload.Sheets(1).Range("A1:O" & LastRowUploadSheet).AutoFilter Field:=15, Criteria1:="#N/A", Operator:=xlOr, Criteria2:="0"
            If PricePointUpload.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                PricePointUpload.Sheets(1).Range("O2:O" & LastRowUploadSheet).SpecialCells(xlCellTypeVisible).ClearContents
                PricePointUpload.Sheets(1).AutoFilter.ShowAllData
            End If
            PricePointUpload.Sheets(1).AutoFilter.ShowAllData
        End If
        PricePointUpload.Sheets(1).AutoFilterMode = False
        ManualSupports.Close False
        PricePointUpload.Close True
        
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=2, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Set ProductDatabase = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Product Database\" & FutureMonthNameDir & "\Product Database.xlsb", ReadOnly:=True)
            ProductDatabase.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
            Master.Sheets("Price File").Range("B12:B" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,2,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").Range("C12:C" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,26,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").Range("D12:D" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,5,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").Range("E12:E" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,4,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").Range("BG12:BG" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,9,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").Range("BH12:BH" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,10,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").Range("BJ12:BJ" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,3,FALSE),""Incorrect WUK Code"")"
            Master.Sheets("Price File").AutoFilter.ShowAllData
            Master.Sheets("Price File").Range("B12:E" & LastRowMaster).Copy
            Master.Sheets("Price File").Range("B12:E" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Price File").Range("BG12:BH" & LastRowMaster).Copy
            Master.Sheets("Price File").Range("BG12:BH" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Price File").Range("BJ12:BJ" & LastRowMaster).Copy
            Master.Sheets("Price File").Range("BJ12:BJ" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            ProductDatabase.Close False
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    
    If LastRowMaster > 11 Then
        Master.Sheets("Price File").Range("I12:I" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[31]-1,0)"
        Master.Sheets("Price File").Range("L12:L" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("O12:O" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("R12:R" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("U12:U" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("X12:X" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("AA12:AA" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("AD12:AD" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("AG12:AG" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("AJ12:AJ" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("AM12:AM" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        Master.Sheets("Price File").Range("AP12:AP" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
        
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=7, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("G12:H" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=10, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("J12:K" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=13, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("M12:N" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=16, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("P12:Q" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=19, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("S12:T" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=22, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("V12:W" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=25, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("Y12:Z" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=28, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("AB12:AC" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=31, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("AE12:AF" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=34, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("AH12:AI" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=37, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("AK12:AL" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=40, Criteria1:="="
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("AN12:AO" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "-"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("BV12:BV" & LastRowMaster).ClearContents
        Master.Sheets("Price File").Range("A12:F" & LastRowMaster).Interior.Color = RGB(208, 206, 206)
        Master.Sheets("Price File").Range("G12:AP" & LastRowMaster).Interior.Color = RGB(155, 194, 230)
        Master.Sheets("Price File").Range("AQ12:AQ" & LastRowMaster).Interior.Color = RGB(226, 239, 218)
        Master.Sheets("Price File").Range("AR12:AX" & LastRowMaster).Interior.Color = RGB(255, 230, 153)
        Master.Sheets("Price File").Range("AY12:BA" & LastRowMaster).Interior.Color = RGB(165, 165, 165)
        Master.Sheets("Price File").Range("BB12:BF" & LastRowMaster).Interior.Color = RGB(255, 192, 0)
        Master.Sheets("Price File").Range("BG12:BM" & LastRowMaster).Interior.Color = RGB(244, 176, 132)
        Master.Sheets("Price File").Range("BN12:BU" & LastRowMaster).Interior.Color = RGB(252, 228, 214)
        Master.Sheets("Price File").Range("BV12:BV" & LastRowMaster).Interior.Color = RGB(198, 224, 180)
        Master.Sheets("Price File").Range("BW11:BX" & LastRowMaster).Interior.Color = RGB(155, 194, 230)
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).Borders.LineStyle = xlContinuous
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).Font.Name = "Calibri"
        Master.Sheets("Price File").Range("A12:BX" & LastRowMaster).Font.Size = 10
        Master.Sheets("Price File").Range("A11:BX11").Font.Size = 11
        Master.Sheets("Price File").Range("G10:BX10").Font.Size = 11
        Master.Sheets("Price File").Range("B12:E" & LastRowMaster).NumberFormat = "General"
        Master.Sheets("Price File").Range("F12:F" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("G12:G" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("J12:J" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("M12:M" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("P12:P" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("S12:S" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("V12:V" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("Y12:Y" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AB12:AB" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AE12:AE" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AH12:AH" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AK12:AK" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AN12:AN" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("H12:I" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("K12:L" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("N12:O" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("Q12:R" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("T12:U" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("W12:X" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("Z12:AA" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AC12:AD" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AF12:AG" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AI12:AJ" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AL12:AM" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AO12:AP" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AQ12:AQ" & LastRowMaster).NumberFormat = "General"
        Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AS12:AS" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AT12:AT" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AU12:AU" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AV12:AW" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("AX12:AX" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("AY12:AY" & LastRowMaster).NumberFormat = "General"
        Master.Sheets("Price File").Range("AZ12:AZ" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BA12:BA" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BB12:BB" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("BC12:BC" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("BD12:BD" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("BE12:BE" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BF12:BF" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BG12:BI" & LastRowMaster).NumberFormat = "General"
        Master.Sheets("Price File").Range("BJ12:BJ" & LastRowMaster).NumberFormat = "General"
        Master.Sheets("Price File").Range("BK12:BK" & LastRowMaster).NumberFormat = "0.00%"
        Master.Sheets("Price File").Range("BL12:BL" & LastRowMaster).NumberFormat = "dd/mm/yyyy"
        Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).NumberFormat = "General"
        Master.Sheets("Price File").Range("BN12:BN" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BO12:BO" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BP12:BP" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BQ12:BQ" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BR12:BR" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BS12:BS" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BT12:BT" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BU12:BU" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BV12:BV" & LastRowMaster).NumberFormat = "$#,##0.00"
        Master.Sheets("Price File").Range("BW11:BX11").Font.Bold = True
        Master.Sheets("Price File").Range("BW11:BX11").Font.Underline = xlUnderlineStyleSingle
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).HorizontalAlignment = xlCenter
        Master.Sheets("Price File").Range("BZ12:BZ" & LastRowMaster).ClearContents
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).VerticalAlignment = xlCenter
        Master.Sheets("Price File").Columns.AutoFit
        Master.Sheets("Price File").Rows.AutoFit
    End If
    Master.Save
End Sub
Sub Price_File_Automation_PostPP()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook
    
    Dim OldPricePoint As Workbook
    Dim NewPricePoint As Workbook
    Dim ProductDatabase As Workbook
    Dim Usage As Workbook
    Dim Spin As Workbook
    
    Master.Sheets("Price File").Columns.EntireColumn.Hidden = False
    Master.Sheets("Price File").Rows.EntireRow.Hidden = False
    
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    LastMonthName = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mmmm")
    CurrentAccount = Master.Sheets("Price File").Range("B3").Value
    Estimator = Master.Sheets("Settings").Range("B1").Value
    SaveNameLoc = Master.Sheets("Price File").Range("B1")
    
    Set NewPricePoint = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\Pricepoint Download\" & FutureMonthNameDir & "\" & CurrentAccount & ".xlsx")
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    LastRowNewPricePoint = NewPricePoint.Sheets(1).Range("H" & Rows.Count).End(xlUp).Row
    
    If Master.Sheets("Price File").AutoFilterMode = True Then
        Master.Sheets("Price File").AutoFilterMode = False
    End If

    Master.Sheets("Price File").Range("A11:AQ" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("A11:AQ" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    NewPricePoint.Sheets(1).Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("Price File").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Master.Sheets("Price File").Range("B7").Value = "01 " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "yyyy")
    Master.Sheets("Price File").Range("AR11").Value = ActiveMonthName & " Terms"
    Master.Sheets("Price File").Range("AS11").Value = ActiveMonthName & " Margin (%)"
    Master.Sheets("Price File").Range("AT11").Value = FutureMonthName & " Terms"
    Master.Sheets("Price File").Range("AU11").Value = FutureMonthName & " Margin (%)"
    Master.Sheets("Price File").Range("AX11").Value = "Above " & FutureMonthName & " Trade"
    Master.Sheets("Price File").Range("AZ11").Value = ActiveMonthName & " Support"
    Master.Sheets("Price File").Range("BA11").Value = FutureMonthName & " Support"
    Master.Sheets("Price File").Range("BN11").Value = ActiveMonthName & " Invoice"
    Master.Sheets("Price File").Range("BO11").Value = FutureMonthName & " Invoice"
    Master.Sheets("Price File").Range("BP11").Value = ActiveMonthName & " Nett Cost"
    Master.Sheets("Price File").Range("BQ11").Value = FutureMonthName & " Nett Cost"
    Master.Sheets("Price File").Range("BR11").Value = ActiveMonthName & " Trade"
    Master.Sheets("Price File").Range("BS11").Value = FutureMonthName & " Trade"
    Master.Sheets("Price File").Range("BT11").Value = ActiveMonthName & " SLP"
    Master.Sheets("Price File").Range("BU11").Value = FutureMonthName & " SLP"
    Master.Sheets("Price File").Range("BV11").Value = "Terms"
    Master.Sheets("Price File").Range("BW11").Value = "Estimator Comments"
    Master.Sheets("Price File").Range("BX11").Value = "Account Manager Comments"
    
    If ActiveMonthName = "January" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-37],2)"
        Else
    End If
        If ActiveMonthName = "February" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-34],2)"
        Else
    End If
        If ActiveMonthName = "March" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-31],2)"
        Else
    End If
        If ActiveMonthName = "April" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-28],2)"
        Else
    End If
        If ActiveMonthName = "May" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-25],2)"
        Else
    End If
        If ActiveMonthName = "June" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-22],2)"
        Else
    End If
        If ActiveMonthName = "July" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-19],2)"
        Else
    End If
        If ActiveMonthName = "August" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-16],2)"
        Else
    End If
        If ActiveMonthName = "September" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-13],2)"
        Else
    End If
        If ActiveMonthName = "October" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-10],2)"
        Else
    End If
        If ActiveMonthName = "November" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-7],2)"
        Else
    End If
        If ActiveMonthName = "December" Then
            Master.Sheets("Price File").Range("AR12:AR" & LastRowMaster).FormulaR1C1 = "=TRUNC(RC[-4],2)"
        Else
    End If
    
    NewPricePoint.Sheets(1).Range("A14:AY" & LastRowNewPricePoint).AutoFilter Field:=20, Criteria1:=">" & Application.EoMonth(Now, 0) + 1
    If NewPricePoint.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        NewPricePoint.Sheets(1).Range("U15:U" & LastRowNewPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        NewPricePoint.Sheets(1).Range("V15:V" & LastRowNewPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        NewPricePoint.Sheets(1).Range("W15:W" & LastRowNewPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        NewPricePoint.Sheets(1).Range("X15:X" & LastRowNewPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        NewPricePoint.Sheets(1).Range("Y15:Y" & LastRowNewPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        NewPricePoint.Sheets(1).Range("AW15:AW" & LastRowNewPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-1]"
        NewPricePoint.Sheets(1).AutoFilter.ShowAllData
        NewPricePoint.Sheets(1).Range("A14:AY" & LastRowNewPricePoint).Copy
        NewPricePoint.Sheets(1).Range("A14:AY" & LastRowNewPricePoint).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    NewPricePoint.Sheets(1).AutoFilter.ShowAllData
    
    Master.Sheets("Price File").Range("AY12:AY" & LastRowMaster).Formula = "=IFERROR(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AJ:$AJ,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))=0,"""",INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AJ:$AJ,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),""No Data"")"
    Master.Sheets("Price File").Range("BA12:BA" & LastRowMaster).Formula = "=IFERROR(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AN:$AN,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))=0,"""",INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AN:$AN,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),""No Data"")"
    Master.Sheets("Price File").Range("BO12:BO" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$W:$W,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$Q:$Q,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BQ12:BQ" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AW:$AW,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AV:$AV,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BS12:BS" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$U:$U,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$O:$O,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BU12:BU" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$X:$X,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$R:$R,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BV12:BV" & LastRowMaster).Formula = "=IFERROR(IF(TRUNC(VLOOKUP(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$AB,21,FALSE),2)=TRUNC(AR12,2),""Terms Match"",TRUNC(VLOOKUP(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$AB,21,FALSE),2)),""No Data"")"
    Master.Sheets("Price File").Range("AS12:BV" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("AS12:BV" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    NewPricePoint.Close False
    
    Set OldPricePoint = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\Pricepoint Download\" & ActiveMonthNameDir & "\" & CurrentAccount & ".xlsx")
    LastRowOldPricePoint = OldPricePoint.Sheets(1).Range("H" & Rows.Count).End(xlUp).Row
    
    OldPricePoint.Sheets(1).Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    OldPricePoint.Sheets(1).Range("A14:AY" & LastRowOldPricePoint).AutoFilter Field:=20, Criteria1:=">" & Application.EoMonth(Now, 0)
    If OldPricePoint.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        OldPricePoint.Sheets(1).Range("U15:U" & LastRowOldPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        OldPricePoint.Sheets(1).Range("V15:V" & LastRowOldPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        OldPricePoint.Sheets(1).Range("W15:W" & LastRowOldPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        OldPricePoint.Sheets(1).Range("X15:X" & LastRowOldPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        OldPricePoint.Sheets(1).Range("Y15:Y" & LastRowOldPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-6]"
        OldPricePoint.Sheets(1).Range("AW15:AW" & LastRowOldPricePoint).SpecialCells(xlCellTypeVisible).Formula = "=RC[-1]"
        OldPricePoint.Sheets(1).AutoFilter.ShowAllData
        OldPricePoint.Sheets(1).Range("A14:AY" & LastRowOldPricePoint).Copy
        OldPricePoint.Sheets(1).Range("A14:AY" & LastRowOldPricePoint).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    OldPricePoint.Sheets(1).AutoFilter.ShowAllData
    
    Master.Sheets("Price File").Range("AZ12:AZ" & LastRowMaster).Formula = "=IFERROR(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AN:$AN,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))=0,"""",INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AN:$AN,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),"""")"
    Master.Sheets("Price File").Range("BN12:BN" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$W:$W,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$Q:$Q,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BP12:BP" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AW:$AW,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AV:$AV,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BR12:BR" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$U:$U,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$O:$O,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("BT12:BT" & LastRowMaster).Formula = "=IFERROR(TRUNC(IF(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$T:$T,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))>0,INDEX([" & CurrentAccount & ".xlsx]Sheet0!$X:$X,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),INDEX([" & CurrentAccount & ".xlsx]Sheet0!$R:$R,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0))),3),""No Data"")"
    Master.Sheets("Price File").Range("AS12:BV" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("AS12:BV" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    OldPricePoint.Close False
    
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=66, Criteria1:="No Data"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Set NewPricePoint = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\Pricepoint Download\" & FutureMonthNameDir & "\" & CurrentAccount & ".xlsx")
        LastRowNewPricePoint = NewPricePoint.Sheets(1).Range("H" & Rows.Count).End(xlUp).Row
        NewPricePoint.Sheets(1).Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Master.Sheets("Price File").Range("BN12:BN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(TRUNC(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$Q:$Q,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),3),""No Data"")"
        Master.Sheets("Price File").Range("BP12:BP" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(TRUNC(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$AV:$AV,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),3),""No Data"")"
        Master.Sheets("Price File").Range("BR12:BR" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(TRUNC(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$O:$O,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),3),""No Data"")"
        Master.Sheets("Price File").Range("BT12:BT" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(TRUNC(INDEX([" & CurrentAccount & ".xlsx]Sheet0!$R:$R,MATCH(@A:A,[" & CurrentAccount & ".xlsx]Sheet0!$H:$H,0)),3),""No Data"")"
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("BN12:BU" & LastRowMaster).Copy
        Master.Sheets("Price File").Range("BN12:BU" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        NewPricePoint.Close False
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    
    Set ProductDatabase = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Product Database\" & FutureMonthNameDir & "\Product Database.xlsb", ReadOnly:=True)
    ProductDatabase.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("Price File").Range("BG12:BG" & LastRowMaster).Formula = "=IFERROR(IF(INDEX('[Product Database.xlsb]Product File (Pyr1)'!$I:$I,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0))=0,"""",INDEX('[Product Database.xlsb]Product File (Pyr1)'!$I:$I,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0))),""Not in Database"")"
    Master.Sheets("Price File").Range("BH12:BH" & LastRowMaster).Formula = "=IFERROR(IF(INDEX('[Product Database.xlsb]Product File (Pyr1)'!$J:$J,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0))=0,"""",INDEX('[Product Database.xlsb]Product File (Pyr1)'!$J:$J,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0))),""Not in Database"")"
    Master.Sheets("Price File").Range("BI12:BI" & LastRowMaster).Formula = "=IFERROR(IF(INDEX('[Product Database.xlsb]Product File (Pyr1)'!$N:$N,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0))=0,"""",INDEX('[Product Database.xlsb]Product File (Pyr1)'!$N:$N,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0))),""Not in Database"")"
    Master.Sheets("Price File").Range("BJ12:BJ" & LastRowMaster).Formula = "=IFERROR(INDEX('[Product Database.xlsb]Product File (Pyr1)'!$C:$C,MATCH(A12,'[Product Database.xlsb]Product File (Pyr1)'!$A:$A,0)),""Missing HoS"")"
    Master.Sheets("Price File").Range("BG12:BJ" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("BG12:BJ" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Master.Sheets("Price File").Range("B12:B" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,2,FALSE),""Incorrect WUK Code"")"
    Master.Sheets("Price File").Range("C12:C" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,26,FALSE),""Incorrect WUK Code"")"
    Master.Sheets("Price File").Range("D12:D" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,5,FALSE),""Incorrect WUK Code"")"
    Master.Sheets("Price File").Range("E12:E" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AA,4,FALSE),""Incorrect WUK Code"")"
    Master.Sheets("Price File").Range("B12:E" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("B12:E" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    ProductDatabase.Close False
    
    If ActiveMonthName = "January" Then
        MAX_FORMULA_RANGE = "MAX(F12,N12,Q12,T12,W12,Z12,AC12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "February" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,Q12,T12,W12,Z12,AC12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "March" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,T12,W12,Z12,AC12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "April" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,W12,Z12,AC12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "May" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,Z12,AC12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "June" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,T12,AC12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "July" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,T12,W12,AF12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "August" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,T12,W12,Z12,AI12,AL12,AO12)"
    End If
    If ActiveMonthName = "September" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,T12,W12,Z12,AC12,AL12,AO12)"
    End If
    If ActiveMonthName = "October" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,T12,W12,Z12,AC12,AF12,AO12)"
    End If
    If ActiveMonthName = "November" Then
        MAX_FORMULA_RANGE = "MAX(F12,H12,K12,N12,Q12,T12,W12,Z12,AC12,AF12,AI12)"
    End If
    If ActiveMonthName = "December" Then
        MAX_FORMULA_RANGE = "MAX(F12,K12,N12,Q12,T12,W12,Z12,AC12,AF12,AI12,AL12)"
    End If
    
    If Master.Sheets("Settings").Range("B7").Value = "MARGIN" Then
        Master.Sheets("Price File").Range("AS12:AS" & LastRowMaster).Formula = "=IFERROR((AR12-BP12)/AR12,""No Data"")"
        Master.Sheets("Price File").Range("AT12:AT" & LastRowMaster).Formula = "=@IFERROR(IF(AND(OR(AV12=""BEST"",AW12=""BEST""),AS12<>""No Data""),BQ12/(1-" & MAX_FORMULA_RANGE & "),IF(OR(AV12=""X"",AW12=""X""),AR12,IF(AND(ISBLANK(AV12),ISBLANK(AW12)),ROUNDUP(BQ12/(1-AS12),2),IF(AND(ISBLANK(AV12),NOT(ISBLANK(AW12))),ROUNDUP(AR12*(1+AW12),2),IF(AND(ISBLANK(AW12),NOT(ISBLANK(AV12))),ROUNDUP(BQ12/(1-AV12),2),""Formula Error""))))),AR:AR)"
        Master.Sheets("Price File").Range("AU12:AU" & LastRowMaster).Formula = "=IFERROR((AT12-BQ12)/AT12,""No Data"")"
        Master.Sheets("Price File").Range("AX12:AX" & LastRowMaster).Formula = "=IFERROR(IF(TRUNC(AT12,2)>TRUNC(BS12,2),TRUNC(AT12,2)-TRUNC(BS12,2),""""),""No Data"")"
        Master.Sheets("Price File").Range("BB12:BB" & LastRowMaster).Formula = "=IFERROR(BU12/BT12-1,""No Data"")"
        Master.Sheets("Price File").Range("BD12:BD" & LastRowMaster).Formula = "=IFERROR(AU12-AS12,""No Data"")"
        Master.Sheets("Price File").Range("BC12:BC" & LastRowMaster).Formula = "=IFERROR(AU12-F12,""No Data"")"
        Master.Sheets("Price File").Range("BE12:BE" & LastRowMaster).Formula = "=IFERROR(AT12-AR12,""No Data"")"
        Master.Sheets("Price File").Range("BF12:BF" & LastRowMaster).Formula = "=IFERROR(BA12-AZ12,BA12)"
    ElseIf Master.Sheets("Settings").Range("B7").Value = "SLP" Then
        Master.Sheets("Price File").Range("AS12:AS" & LastRowMaster).Formula = "=IFERROR((AR12-BP12)/AR12,""No Data"")"
        Master.Sheets("Price File").Range("AT12:AT" & LastRowMaster).Formula = "=@IFERROR(IF(AND(OR(AV12=""BEST"",AW12=""BEST""),AS12<>""No Data""),BQ12/(1-" & MAX_FORMULA_RANGE & "),IF(OR(AV12=""X"",AW12=""X""),AR12,IF(AND(ISBLANK(AV12),ISBLANK(AW12)),ROUNDUP(AR12*(1+BB12),2),IF(AND(ISBLANK(AV12),NOT(ISBLANK(AW12))),ROUNDUP(AR12*(1+AW12),2),IF(AND(ISBLANK(AW12),NOT(ISBLANK(AV12))),ROUNDUP(BQ12/(1-AV12),2),""Formula Error""))))),AR:AR)"
        Master.Sheets("Price File").Range("AU12:AU" & LastRowMaster).Formula = "=IFERROR((AT12-BQ12)/AT12,""No Data"")"
        Master.Sheets("Price File").Range("AX12:AX" & LastRowMaster).Formula = "=IFERROR(IF(TRUNC(AT12,2)>TRUNC(BS12,2),TRUNC(AT12,2)-TRUNC(BS12,2),""""),""No Data"")"
        Master.Sheets("Price File").Range("BB12:BB" & LastRowMaster).Formula = "=IFERROR(BU12/BT12-1,""No Data"")"
        Master.Sheets("Price File").Range("BD12:BD" & LastRowMaster).Formula = "=IFERROR(AU12-AS12,""No Data"")"
        Master.Sheets("Price File").Range("BC12:BC" & LastRowMaster).Formula = "=IFERROR(AU12-F12,""No Data"")"
        Master.Sheets("Price File").Range("BE12:BE" & LastRowMaster).Formula = "=IFERROR(AT12-AR12,""No Data"")"
        Master.Sheets("Price File").Range("BF12:BF" & LastRowMaster).Formula = "=IFERROR(BA12-AZ12,BA12)"
    End If
    
    If ActiveMonthName = "January" Then
        SELL_PRICE_FUTUREMONTH = "J12:J"
        MARGIN_FUTUREMONTH = "K12:K"
        MARGIN_CURRENTMONTH = "H12:H"
    End If
    If ActiveMonthName = "February" Then
        SELL_PRICE_FUTUREMONTH = "M12:M"
        MARGIN_FUTUREMONTH = "N12:N"
        MARGIN_CURRENTMONTH = "K12:K"
    End If
    If ActiveMonthName = "March" Then
        SELL_PRICE_FUTUREMONTH = "P12:P"
        MARGIN_FUTUREMONTH = "Q12:Q"
        MARGIN_CURRENTMONTH = "N12:N"
    End If
    If ActiveMonthName = "April" Then
        SELL_PRICE_FUTUREMONTH = "S12:S"
        MARGIN_FUTUREMONTH = "T12:T"
        MARGIN_CURRENTMONTH = "Q12:Q"
    End If
    If ActiveMonthName = "May" Then
        SELL_PRICE_FUTUREMONTH = "V12:V"
        MARGIN_FUTUREMONTH = "W12:W"
        MARGIN_CURRENTMONTH = "T12:T"
    End If
    If ActiveMonthName = "June" Then
        SELL_PRICE_FUTUREMONTH = "Y12:Y"
        MARGIN_FUTUREMONTH = "Z12:Z"
        MARGIN_CURRENTMONTH = "W12:W"
    End If
    If ActiveMonthName = "July" Then
        SELL_PRICE_FUTUREMONTH = "AB12:AB"
        MARGIN_FUTUREMONTH = "AC12:AC"
        MARGIN_CURRENTMONTH = "Z12:Z"
    End If
    If ActiveMonthName = "August" Then
        SELL_PRICE_FUTUREMONTH = "AE12:AE"
        MARGIN_FUTUREMONTH = "AF12:AF"
        MARGIN_CURRENTMONTH = "AC12:AC"
    End If
    If ActiveMonthName = "September" Then
        SELL_PRICE_FUTUREMONTH = "AH12:AH"
        MARGIN_FUTUREMONTH = "AI12:AI"
        MARGIN_CURRENTMONTH = "AF12:AF"
    End If
    If ActiveMonthName = "October" Then
        SELL_PRICE_FUTUREMONTH = "AK12:AK"
        MARGIN_FUTUREMONTH = "AL12:AL"
        MARGIN_CURRENTMONTH = "AI12:AI"
    End If
    If ActiveMonthName = "November" Then
        SELL_PRICE_FUTUREMONTH = "AN12:AN"
        MARGIN_FUTUREMONTH = "AO12:AO"
        MARGIN_CURRENTMONTH = "AL12:AL"
    End If
    If ActiveMonthName = "December" Then
        SELL_PRICE_FUTUREMONTH = "G12:G"
        MARGIN_FUTUREMONTH = "H12:H"
        MARGIN_CURRENTMONTH = "AO12:AO"
    End If
    
    Master.Sheets("Price File").Range(SELL_PRICE_FUTUREMONTH & LastRowMaster).Formula = "=TRUNC(AT:AT,2)"
    Master.Sheets("Price File").Range(MARGIN_FUTUREMONTH & LastRowMaster).Formula = "=@AU:AU"
    Master.Sheets("Price File").Range(MARGIN_CURRENTMONTH & LastRowMaster).Formula = "=@AS:AS"
    Master.Sheets("Price File").Range("I12:I" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[31]-1,0)"
    Master.Sheets("Price File").Range("L12:L" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("O12:O" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("R12:R" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("U12:U" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("X12:X" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("AA12:AA" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("AD12:AD" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("AG12:AG" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("AJ12:AJ" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("AM12:AM" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    Master.Sheets("Price File").Range("AP12:AP" & LastRowMaster).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
    
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=7, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("G12:G" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("H12:H" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("I12:I" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=10, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("J12:J" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("K12:K" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("L12:L" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=13, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("M12:M" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("N12:N" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("O12:O" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=16, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("P12:P" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("Q12:Q" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("R12:R" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=19, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("S12:S" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("T12:T" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("U12:U" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=22, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("V12:V" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("W12:W" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("X12:X" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=25, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("Y12:Y" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("Z12:Z" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AA12:AA" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=28, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AB12:AB" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AC12:AC" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AD12:AD" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=31, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AE12:AE" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AF12:AF" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AG12:AG" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=34, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AH12:AH" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AI12:AI" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AJ12:AJ" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=37, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AK12:AK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AL12:AL" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AM12:AM" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=40, Criteria1:="="
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AN12:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AO12:AO" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).Range("AP12:AP" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "-"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=9, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("I12:I" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[31]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=12, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("L12:L" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=15, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("O12:O" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=18, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("R12:R" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=21, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("U12:U" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=24, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("X12:X" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=27, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AA12:AA" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=30, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AD12:AD" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=33, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets(1).Range("AG12:AG" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=36, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 12 Then
            Master.Sheets(1).Range("AJ12:AJ" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=39, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 12 Then
            Master.Sheets(1).Range("AM12:AM" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=1, Criteria1:="<>"
    Master.Sheets(1).Range("A11:BX" & LastRowMaster).AutoFilter Field:=42, Criteria1:="<>"
        If Master.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 12 Then
            Master.Sheets(1).Range("AP12:AP" & LastRowMaster).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-5]-1,0)"
            Master.Sheets(1).AutoFilter.ShowAllData
            Else
        Master.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets(1).AutoFilter.ShowAllData
    
    Set GetExpiredSupports = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\PricePoint Download\" & ActiveMonthNameDir & "\" & CurrentAccount & ".xlsx")
    LastRowGetExpiredSupport = GetExpiredSupports.Sheets(1).Range("H" & Rows.Count).End(xlUp).Row
    
    GetExpiredSupports.Sheets(1).Range("A14:AY" & LastRowGetExpiredSupport).AutoFilter Field:=37, Criteria1:="<=" & Application.EoMonth(Now, -1)
        If GetExpiredSupports.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets.Add(After:=Sheets(1)).Name = "Expired"
            GetExpiredSupports.Sheets(1).Range("H14:I" & LastRowGetExpiredSupport).SpecialCells(xlCellTypeVisible).Copy
            GetExpiredSupports.Sheets("Expired").Range("A1").SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            GetExpiredSupports.Sheets(1).Range("D14:E" & LastRowGetExpiredSupport).SpecialCells(xlCellTypeVisible).Copy
            GetExpiredSupports.Sheets("Expired").Range("C1").SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            GetExpiredSupports.Sheets(1).Range("AJ14:AK" & LastRowGetExpiredSupport).SpecialCells(xlCellTypeVisible).Copy
            GetExpiredSupports.Sheets("Expired").Range("E1").SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            GetExpiredSupports.Sheets("Expired").Columns("F:F").NumberFormat = "dd/mm/yyyy"
            GetExpiredSupports.Sheets("Expired").Range("A1:F" & LastRowGetExpiredSupport).Columns.AutoFit
            GetExpiredSupports.Sheets("Expired").Copy
            ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Expired Supports\" & FutureMonthNameDir & "\" & CurrentAccount & " - " & SaveNameLoc & " - Expired Supports.xlsx"
            ActiveWorkbook.Close False
            Application.CutCopyMode = False
            Else
        GetExpiredSupports.Sheets(1).AutoFilter.ShowAllData
    End If
    GetExpiredSupports.Close False
    
    Set GetExpiredSupports = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\PricePoint Download\" & ActiveMonthNameDir & "\" & CurrentAccount & ".xlsx")
    LastRowGetExpiredSupport = GetExpiredSupports.Sheets(1).Range("H" & Rows.Count).End(xlUp).Row
    
    GetExpiredSupports.Sheets(1).Range("A14:AY" & LastRowGetExpiredSupport).AutoFilter Field:=37, Criteria1:="<=" & Application.EoMonth(Now, 0), Operator:=xlAnd, Criteria2:=">" & Application.EoMonth(Now, -1)
        If GetExpiredSupports.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets.Add(After:=Sheets(1)).Name = "Expired"
            GetExpiredSupports.Sheets(1).Range("H14:I" & LastRowGetExpiredSupport).SpecialCells(xlCellTypeVisible).Copy
            GetExpiredSupports.Sheets("Expired").Range("A1").SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            GetExpiredSupports.Sheets(1).Range("D14:E" & LastRowGetExpiredSupport).SpecialCells(xlCellTypeVisible).Copy
            GetExpiredSupports.Sheets("Expired").Range("C1").SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            GetExpiredSupports.Sheets(1).Range("AJ14:AK" & LastRowGetExpiredSupport).SpecialCells(xlCellTypeVisible).Copy
            GetExpiredSupports.Sheets("Expired").Range("E1").SpecialCells(xlCellTypeVisible).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            GetExpiredSupports.Sheets("Expired").Columns("F:F").NumberFormat = "dd/mm/yyyy"
            GetExpiredSupports.Sheets("Expired").Range("A1:F" & LastRowGetExpiredSupport).Columns.AutoFit
            GetExpiredSupports.Sheets("Expired").Copy
            ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\End of Month Supports\" & FutureMonthNameDir & "\" & CurrentAccount & " - " & SaveNameLoc & " - About to Expire.xlsx"
            ActiveWorkbook.Close False
            Application.CutCopyMode = False
            Else
        GetExpiredSupports.Sheets(1).AutoFilter.ShowAllData
    End If
    GetExpiredSupports.Close False
    
    Set Spin = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Spin File\" & FutureMonthNameDir & "\SPIN.xlsx", ReadOnly:=True)
    LastRowSpin = Spin.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterNextMonth, Operator:=xlFilterDynamic
        If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Next Month"
                Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
                Spin.Sheets("Next Month").Range("A1").PasteSpecial Paste:=xlPasteValues
                Spin.Sheets("Next Month").Range("A1").PasteSpecial Paste:=xlFormats
                Application.CutCopyMode = False
                Spin.Sheets(1).AutoFilter.ShowAllData
            Else
        Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterThisMonth, Operator:=xlFilterDynamic
    If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "This Month"
            Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
            Spin.Sheets("This Month").Range("A1").PasteSpecial Paste:=xlPasteValues
            Spin.Sheets("This Month").Range("A1").PasteSpecial Paste:=xlFormats
            Application.CutCopyMode = False
            Spin.Sheets(1).AutoFilter.ShowAllData
        Else
    Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterLastMonth, Operator:=xlFilterDynamic
    If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Last Month"
            Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
            Spin.Sheets("Last Month").Range("A1").PasteSpecial Paste:=xlPasteValues
            Spin.Sheets("Last Month").Range("A1").PasteSpecial Paste:=xlFormats
            Application.CutCopyMode = False
            Spin.Sheets(1).AutoFilter.ShowAllData
        Else
    Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).Formula = "=IFERROR(IFERROR(IFERROR(IF(VLOOKUP(BJ12,'[SPIN.xlsx]Next Month'!$C:$C,1,FALSE)=BJ12,""Next Month"",""Not inline""),IF(VLOOKUP(BJ12,'[SPIN.xlsx]This Month'!$C:$C,1,FALSE)=BJ12,""This Month"",""Not inline"")),IF(VLOOKUP(BJ12,'[SPIN.xlsx]Last Month'!$C:$C,1,FALSE)=BJ12,""Last Month"",""Not inline"")),""Not on Spin"")"
    Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:="This Month"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = ActiveMonthName & " Increase"
            Master.Sheets("Price File").Range("BK12:BK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]This Month'!$C:$J,4,FALSE)"
            Master.Sheets("Price File").Range("BL12:BL" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]This Month'!$C:$J,5,FALSE)"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:="Last Month"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = LastMonthName & " Increase"
            Master.Sheets("Price File").Range("BK12:BK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]Last Month'!$C:$J,4,FALSE)"
            Master.Sheets("Price File").Range("BL12:BL" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]Last Month'!$C:$J,5,FALSE)"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:="Next Month"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = FutureMonthName & " Increase"
            Master.Sheets("Price File").Range("BK12:BK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]Next Month'!$C:$J,4,FALSE)"
            Master.Sheets("Price File").Range("BL12:BL" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]Next Month'!$C:$J,5,FALSE)"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("BK12:BM" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("BK12:BM" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Spin.Close False
    
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:="Not on Spin"
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=54, Criteria1:=">0%"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BM12:BM" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "Not Inline Increase"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    
    If Master.Sheets("Settings").Range("B6").Value = "YES" Then
        Set Usage = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\DATA\Usage\" & FutureMonthNameDir & "\" & CurrentAccount & ".xlsx")
        Usage.Sheets(1).Rows(1).EntireRow.Delete
        LastRowUsage = Usage.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Usage.Sheets(1).Columns(1).TextToColumns Destination:=Columns(1), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Master.Sheets("Price File").Range("AQ12:AQ" & LastRowMaster).Formula = "=IFERROR(IF(VLOOKUP(@A:A,[" & CurrentAccount & ".xlsx]Quantity!$A:$C,3,FALSE)<0,""No Usage"",VLOOKUP(@A:A,[" & CurrentAccount & ".xlsx]Quantity!$A:$C,3,FALSE)),""No Usage"")"
        Master.Sheets("Price File").Range("AQ12:AQ" & LastRowMaster).Copy
        Master.Sheets("Price File").Range("AQ12:AQ" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Usage.Close False
    Else
        Master.Sheets("Price File").Range("AQ12:AQ" & LastRowMaster).Value = "No Data Available"
    End If
    
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=6, Criteria1:="0.00%", Operator:=xlOr, Criteria2:="="
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets("Price File").Range("F12:F" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=@AS:AS"
        Master.Sheets("Price File").AutoFilter.ShowAllData
        Master.Sheets("Price File").Range("F12:F" & LastRowMaster).Copy
        Master.Sheets("Price File").Range("F12:F" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").AutoFilter.ShowAllData
    Else
        Master.Sheets("Price File").AutoFilter.ShowAllData
    End If

    Master.Sheets("Price File").Columns.AutoFit
    Master.Sheets("Price File").Rows.AutoFit
        If ActiveMonthName = "January" Then
        Master.Sheets("Price File").Columns("M:AP").Hidden = True
        Else
    End If
        If ActiveMonthName = "February" Then
        Master.Sheets("Price File").Columns("P:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:I").Hidden = True
        Else
    End If
        If ActiveMonthName = "March" Then
        Master.Sheets("Price File").Columns("S:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:L").Hidden = True
        Else
    End If
        If ActiveMonthName = "April" Then
        Master.Sheets("Price File").Columns("V:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:O").Hidden = True
        Else
    End If
        If ActiveMonthName = "May" Then
        Master.Sheets("Price File").Columns("Y:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:R").Hidden = True
        Else
    End If
        If ActiveMonthName = "June" Then
        Master.Sheets("Price File").Columns("AB:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:U").Hidden = True
        Else
    End If
        If ActiveMonthName = "July" Then
        Master.Sheets("Price File").Columns("AE:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:X").Hidden = True
        Else
    End If
        If ActiveMonthName = "August" Then
        Master.Sheets("Price File").Columns("AH:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:AA").Hidden = True
        Else
    End If
        If ActiveMonthName = "September" Then
        Master.Sheets("Price File").Columns("AK:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:AD").Hidden = True
        Else
    End If
        If ActiveMonthName = "October" Then
        Master.Sheets("Price File").Columns("AN:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:AD").Hidden = True
        Else
    End If
        If ActiveMonthName = "November" Then
        Master.Sheets("Price File").Columns("G:AJ").Hidden = True
        Else
    End If
        If ActiveMonthName = "December" Then
        Master.Sheets("Price File").Columns("J:AM").Hidden = True
        Else
    End If
    
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=59, Criteria1:="EOL"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Obsolete"
        Master.Sheets("Price File").Range("A11:E" & LastRowMaster).Copy
        Master.Sheets("Obsolete").Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").Range("BG11:BI" & LastRowMaster).Copy
        Master.Sheets("Obsolete").Range("F1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").Range("AQ11:AQ" & LastRowMaster).Copy
        Master.Sheets("Obsolete").Range("I1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowObsolete = Master.Sheets("Obsolete").Range("A" & Rows.Count).End(xlUp).Row
        Master.Sheets("Obsolete").Range("J1").Value = "New WUK Code"
        Master.Sheets("Obsolete").Range("K1").Value = "New Description"
        Master.Sheets("Obsolete").Range("L1").Value = "New SPG Group"
        Master.Sheets("Obsolete").Range("M1").Value = "New Manufacturer Code"
        Master.Sheets("Obsolete").Range("N1").Value = "New Supplier Name"
        Master.Sheets("Obsolete").Range("O1").Value = "New LCC"
        Master.Sheets("Obsolete").Range("P1").Value = "New Product Narrative"
        Master.Sheets("Obsolete").Range("Q1").Value = "Comment Box"
        Master.Sheets("Obsolete").Range("A1:E" & LastRowObsolete).Interior.Color = RGB(208, 206, 206)
        Master.Sheets("Obsolete").Range("F1:I" & LastRowObsolete).Interior.Color = RGB(226, 239, 218)
        Master.Sheets("Obsolete").Range("J1:Q" & LastRowObsolete).Interior.Color = RGB(208, 206, 206)
        Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).Borders.LineStyle = xlContinuous
        Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).Font.Name = "Calibri"
        Master.Sheets("Obsolete").Range("A2:Q" & LastRowObsolete).Font.Size = 10
        Master.Sheets("Obsolete").Range("A1:Q1").Font.Size = 11
        Master.Sheets("Obsolete").Range("A1:Q1").Font.Bold = True
        Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).AutoFilter
        Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).HorizontalAlignment = xlCenter
        Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).VerticalAlignment = xlCenter
        Master.Sheets("Obsolete").Range("J2:J" & LastRowObsolete).Formula = "=IF(ISNUMBER(SEARCH(""replaces"",@H:H))=FALSE,WUKCODE(@H:H),""Unknown Code"")"
        Master.Sheets("Obsolete").Range("J2:J" & LastRowObsolete).Copy
        Master.Sheets("Obsolete").Range("J2:J" & LastRowObsolete).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Obsolete").Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Master.Sheets("Obsolete").Range("Q2:Q" & LastRowObsolete).Formula = "=IFERROR(IF(VLOOKUP(@J:J,'Price File'!A:A,1,FALSE)=@J:J,""New Code Already on file"",""""),"""")"
        Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).AutoFilter Field:=10, Criteria1:="<>Unknown Code"
        If Master.Sheets("Obsolete").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Set ProductDatabase = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Product Database\" & FutureMonthNameDir & "\Product Database.xlsb", ReadOnly:=True)
            ProductDatabase.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
            Master.Sheets("Obsolete").Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
            Master.Sheets("Obsolete").Range("K2:K" & LastRowObsolete).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@J:J,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AD,2,FALSE),""New Code Not in Database"")"
            Master.Sheets("Obsolete").Range("L2:L" & LastRowObsolete).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@J:J,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AD,26,FALSE),""New Code Not in Database"")"
            Master.Sheets("Obsolete").Range("M2:M" & LastRowObsolete).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@J:J,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AD,5,FALSE),""New Code Not in Database"")"
            Master.Sheets("Obsolete").Range("N2:N" & LastRowObsolete).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@J:J,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AD,4,FALSE),""New Code Not in Database"")"
            Master.Sheets("Obsolete").Range("O2:O" & LastRowObsolete).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@J:J,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AD,9,FALSE),""New Code Not in Database"")"
            Master.Sheets("Obsolete").Range("P2:P" & LastRowObsolete).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@J:J,'[Product Database.xlsb]Product File (Pyr1)'!$A:$AD,14,FALSE),""New Code Not in Database"")"
            Master.Sheets("Obsolete").AutoFilter.ShowAllData
            Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).Copy
            Master.Sheets("Obsolete").Range("A1:Q" & LastRowObsolete).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            ProductDatabase.Close False
        End If
        Application.CutCopyMode = False
        Master.Sheets("Obsolete").Columns.AutoFit
        Master.Sheets("Obsolete").Rows.AutoFit
        Master.Sheets("Obsolete").Activate
        Master.Sheets("Obsolete").Copy
        ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Obsoletes\" & FutureMonthNameDir & "\" & CurrentAccount & " - " & SaveNameLoc & " - Obsoletes.xlsx"
        ActiveWorkbook.Close True
        Master.Sheets("Obsolete").Delete
        If Master.Sheets("Settings").Range("A8").Value <> "Auto EOL Removal?" Then
            Master.Sheets("Settings").Range("A8").Value = "Auto EOL Removal?"
            Master.Sheets("Settings").Range("B8").Value = "NO"
        End If
        If Master.Sheets("Settings").Range("B8").Value = "YES" Then
            Master.Sheets("Price File").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
    Else
        Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").Activate
    If Master.Sheets("Settings").Range("B6").Value = "YES" Then
        Dim exists As Boolean
        For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Basket" Then
            exists = True
        End If
        Next i
        If Not exists Then
            Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "Basket"
        End If
        Master.Sheets("Price File").Range("AR2").Value = "Total Sell:"
        Master.Sheets("Price File").Range("AR3").Value = "Total Cost:"
        Master.Sheets("Price File").Range("AR4").Value = "Basket Margin:"
        Master.Sheets("Price File").Range("AR5").Value = "Total Profit:"
        Master.Sheets("Price File").Range("AT2").Value = "Total Sell:"
        Master.Sheets("Price File").Range("AT3").Value = "Total Cost:"
        Master.Sheets("Price File").Range("AT4").Value = "Basket Margin:"
        Master.Sheets("Price File").Range("AT5").Value = "Total Profit:"
        Master.Sheets("Price File").Range("AR1:AS1").Merge
        Master.Sheets("Price File").Range("AT1:AU1").Merge
        Master.Sheets("Price File").Range("AR1:AV1,AR2:AR5,AT2:AT5").Interior.Color = RGB(51, 204, 255)
        Master.Sheets("Price File").Range("AR1:AV5").Borders.LineStyle = xlContinuous
        Master.Sheets("Price File").Range("AR1:AV1").HorizontalAlignment = xlCenter
        Master.Sheets("Price File").Range("AR1:AV1").VerticalAlignment = xlCenter
        Master.Sheets("Price File").Range("BY11:CA" & LastRowMaster).ClearContents
        Master.Sheets("Price File").Range("AR1").Value = ActiveMonthName & " Basket"
        Master.Sheets("Price File").Range("AT1").Value = FutureMonthName & " Basket"
        Master.Sheets("Price File").Range("AV1").Value = ActiveMonthName & " Vs. " & FutureMonthName
        Master.Sheets("Price File").Range("A11:CA" & LastRowMaster).AutoFilter Field:=43, Criteria1:=">0%", Operator:=xlAnd, Criteria2:="<>No Usage"
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("basket").Cells.Clear
            Master.Sheets("Price File").Range("A12:A" & LastRowMaster).Copy
            Master.Sheets("Basket").Range("A1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowBasket = Master.Sheets("Basket").Range("A" & Rows.Count).End(xlUp).Row
            Master.Sheets("Basket").Range("B1:B" & LastRowBasket).Formula = "=IFERROR(VLOOKUP(@A:A,'Price File'!A:BU,68,FALSE)*VLOOKUP(@A:A,'Price File'!A:BU,43,FALSE),0)"
            Master.Sheets("Basket").Range("C1:C" & LastRowBasket).Formula = "=IFERROR(VLOOKUP(@A:A,'Price File'!A:BU,44,FALSE)*VLOOKUP(@A:A,'Price File'!A:BU,43,FALSE),0)"
            Master.Sheets("Basket").Range("D1:D" & LastRowBasket).Formula = "=IFERROR(VLOOKUP(@A:A,'Price File'!A:BU,69,FALSE)*VLOOKUP(@A:A,'Price File'!A:BU,43,FALSE),0)"
            Master.Sheets("Basket").Range("E1:E" & LastRowBasket).Formula = "=IFERROR(VLOOKUP(@A:A,'Price File'!A:BU,46,FALSE)*VLOOKUP(@A:A,'Price File'!A:BU,43,FALSE),0)"
            Master.Sheets("Price File").Range("AS2:AV5").NumberFormat = "$#,##0.00"
            Master.Sheets("Price File").Range("AS4:AV4").NumberFormat = "0.00%"
            Master.Sheets("Price File").Range("AS4").Formula = "=(AS2-AS3)/AS2"
            Master.Sheets("Price File").Range("AS2").Formula = "=SUM(Basket!C:C)"
            Master.Sheets("Price File").Range("AS3").Formula = "=SUM(Basket!B:B)"
            Master.Sheets("Price File").Range("AS5").Formula = "=AS2-AS3"
            Master.Sheets("Price File").Range("AU4").Formula = "=(AU2-AU3)/AU2"
            Master.Sheets("Price File").Range("AU2").Formula = "=SUM(Basket!E:E)"
            Master.Sheets("Price File").Range("AU3").Formula = "=SUM(Basket!D:D)"
            Master.Sheets("Price File").Range("AU5").Formula = "=AU2-AU3"
            Master.Sheets("Price File").Range("AV2").Formula = "=AU2-AS2"
            Master.Sheets("Price File").Range("AV3").Formula = "=AU3-AS3"
            Master.Sheets("Price File").Range("AV4").Formula = "=AU4-AS4"
            Master.Sheets("Price File").Range("AV5").Formula = "=AU5-AS5"
            Application.CutCopyMode = False
        End If
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Save
End Sub
Sub Export_Terms()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.Zoom = 85
    ActiveWindow.DisplayGridlines = False
    
    Dim Master As ThisWorkbook
    Dim Additions As Workbook
    Dim Estimator As String
    Dim CurrentAccount As String
    Dim CurrentAccountTerms As String
    Dim ActiveMonthName As String
    Dim PricePointUpload As Workbook
    Dim answer As Integer
    Dim ExpiryDate As String
    Dim ProductDatabase As Workbook
    
    Set Master = ThisWorkbook
    
    Master.Sheets("Price File").Columns.EntireColumn.Hidden = False
    Master.Sheets("Price File").Rows.EntireRow.Hidden = False
    
    Estimator = Master.Sheets("Settings").Range("B1").Value
    CurrentAccount = Master.Sheets("Price File").Range("B3").Value
    If Master.Sheets("Settings").Range("B10").Value = "" Then
        CurrentAccountTerms = Master.Sheets("Price File").Range("B3").Value
    Else
        CurrentAccountTerms = Master.Sheets("Settings").Range("B10").Value
    End If
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    Branchcode = Master.Sheets("Settings").Range("B4").Value
    AccountFullName = Master.Sheets("Price File").Range("B1").Value
    ContractName = Master.Sheets("Settings").Range("B3").Value
    ExpiryDate = Master.Sheets("Settings").Range("B5").Value
    Master.Activate
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=48, Criteria1:="*rem*"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets.Add.Name = "REMOVALS"
        Master.Sheets("Price File").Range("A11:E" & LastRowMaster).Copy
        Master.Sheets("REMOVALS").Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").Range("AR11:AR" & LastRowMaster).Copy
        Master.Sheets("REMOVALS").Range("F1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("REMOVALS").Columns.AutoFit
        Master.Sheets("REMOVALS").Rows.AutoFit
        Master.Sheets("Price File").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
        If Master.Sheets("Price File").AutoFilterMode = True Then Master.Sheets("Price File").AutoFilterMode = False
        Master.Sheets("REMOVALS").Copy
        ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Terms\" & FutureMonthNameDir & "\" & CurrentAccount & " " & AccountFullName & " " & FutureMonthName & " - Terms Removal.csv", FileFormat:=xlCSV
        ActiveWorkbook.Close True
        Master.Activate
        Master.Sheets("REMOVALS").Delete
        Master.Save
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    
    
    Set ProductDatabase = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Product Database\" & FutureMonthNameDir & "\Product Database.xlsb", ReadOnly:=True)
    ProductDatabase.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("Price File").Range("BG12:BG" & LastRowMaster).Formula = "=IFERROR(VLOOKUP(A12,'[product Database.xlsb]Product File (Pyr1)'!$A:$I,9,FALSE),""No Data"")"
    Master.Sheets("Price File").Range("BG12:BG" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("BG12:BG" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    ProductDatabase.Close False
    
    
    If Master.Sheets("settings").Range("B2").Value = "FPE" Then
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=59, Criteria1:=Array("BPO", "DCS", "STX", "STO", "DCD", "STC", "NFX", "DCX", "REG", "NEW", "CHG", "NWR", "NYA", "CRE", "XDK", "AQA", "TGM", "DIR", "PRV", "HOL"), Operator:=xlFilterValues
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=3, Criteria1:="<>"
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        
        Master.Sheets.Add(After:=Sheets("Settings")).Name = "Terms"
        Master.Sheets("Terms").Range("A1").Value = "Q"
        Master.Sheets("Terms").Range("B1").Value = "Quote"
        Master.Sheets("Terms").Range("C1").Value = Branchcode & "/" & CurrentAccountTerms
        Master.Sheets("Terms").Range("D1").Value = "Customer - " & CurrentAccountTerms
        Master.Sheets("Terms").Range("E1").Value = CurrentAccountTerms
        Master.Sheets("Terms").Range("F1").Value = "q0"
        Master.Sheets("Terms").Range("G1").Value = "39352"
        Master.Sheets("Terms").Range("H1").Value = "14:45:46"
        Master.Sheets("Terms").Range("I1").Value = "almir"
        
        Master.Sheets("Terms").Range("A2").Value = "Item"
        Master.Sheets("Terms").Range("B2").Value = "Description"
        Master.Sheets("Terms").Range("C2").Value = "Discount 1"
        Master.Sheets("Terms").Range("D2").Value = "Discount 2"
        Master.Sheets("Terms").Range("E2").Value = "Quantity"
        Master.Sheets("Terms").Range("F2").Value = "Per(Qty)"
        Master.Sheets("Terms").Range("G2").Value = "Price"
        Master.Sheets("Terms").Range("H2").Value = "Per(Price)"
        
        Master.Sheets("Price File").Range("A12:A" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Terms").Range("A3").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
            If ActiveMonthName = "January" Then
                Master.Sheets("Price File").Range("J12:J" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "February" Then
                Master.Sheets("Price File").Range("M12:M" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "March" Then
                Master.Sheets("Price File").Range("P12:P" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "April" Then
                Master.Sheets("Price File").Range("S12:S" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "May" Then
                Master.Sheets("Price File").Range("V12:V" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "June" Then
                Master.Sheets("Price File").Range("Y12:Y" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "July" Then
                Master.Sheets("Price File").Range("AB12:AB" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "August" Then
                Master.Sheets("Price File").Range("AE12:AE" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "September" Then
                Master.Sheets("Price File").Range("AH12:AH" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "October" Then
                Master.Sheets("Price File").Range("AK12:AK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "November" Then
                Master.Sheets("Price File").Range("AN12:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "December" Then
                Master.Sheets("Price File").Range("G12:G" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            Else
        Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        LastRowTerms = Master.Sheets("Terms").Range("A" & Rows.Count).End(xlUp).Row
        Master.Sheets("Terms").Range("E3:E" & LastRowTerms).Value = "1"
        If Master.Sheets("Settings").Range("B2").Value = "FPE" Then
            Master.Sheets("Terms").Copy
            ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Terms\" & FutureMonthNameDir & "\" & Branchcode & " - " & CurrentAccount & " " & AccountFullName & " " & FutureMonthName & " - FPE Terms.csv", FileFormat:=xlCSV
            ActiveWorkbook.Close True
        End If
    Else
    End If
    
    If Master.Sheets("Settings").Range("B2").Value = "CONTRACT" Then
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=59, Criteria1:=Array("BPO", "DCS", "STX", "STO", "DCD", "STC", "NFX", "DCX", "REG", "NEW", "CHG", "NWR", "NYA", "CRE", "XDK", "AQA", "TGM", "DIR", "PRV", "HOL"), Operator:=xlFilterValues
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=3, Criteria1:="<>"
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        
        Master.Sheets.Add(After:=Sheets("Settings")).Name = "Terms"
        Master.Sheets("Price File").Range("A12:A" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Terms").Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowTerms = Master.Sheets("Terms").Range("A" & Rows.Count).End(xlUp).Row
        Master.Sheets("Price File").Range("C12:C" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Terms").Range("D1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").Range("B12:B" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Terms").Range("B1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
            If ActiveMonthName = "January" Then
                Master.Sheets("Price File").Range("J12:J" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "February" Then
                Master.Sheets("Price File").Range("M12:M" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "March" Then
                Master.Sheets("Price File").Range("P12:P" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "April" Then
                Master.Sheets("Price File").Range("S12:S" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "May" Then
                Master.Sheets("Price File").Range("V12:V" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "June" Then
                Master.Sheets("Price File").Range("Y12:Y" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "July" Then
                Master.Sheets("Price File").Range("AB12:AB" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "August" Then
                Master.Sheets("Price File").Range("AE12:AE" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "September" Then
                Master.Sheets("Price File").Range("AH12:AH" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "October" Then
                Master.Sheets("Price File").Range("AK12:AK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "November" Then
                Master.Sheets("Price File").Range("AN12:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "December" Then
                Master.Sheets("Price File").Range("G12:G" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
        If Master.Sheets("Settings").Range("B2").Value = "CONTRACT" Then
            Master.Sheets("Terms").Copy
            ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Terms\" & FutureMonthNameDir & "\" & ContractName & " - " & AccountFullName & " " & FutureMonthName & " - Contract Terms.xlsx"
            ActiveWorkbook.Close True
        End If
            Else
        Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    
    If Master.Sheets("Settings").Range("B2").Value = "BOTH" Then
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=59, Criteria1:=Array("BPO", "DCS", "STX", "STO", "DCD", "STC", "NFX", "DCX", "REG", "NEW", "CHG", "NWR", "NYA", "CRE", "XDK", "AQA", "TGM", "DIR", "PRV", "HOL"), Operator:=xlFilterValues
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=3, Criteria1:="<>"
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets.Add(After:=Sheets("Settings")).Name = "Contract Terms"
        Master.Sheets("Price File").Range("A12:A" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Contract Terms").Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowTerms = Master.Sheets("Contract Terms").Range("A" & Rows.Count).End(xlUp).Row
        Master.Sheets("Price File").Range("C12:C" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Contract Terms").Range("D1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").Range("B12:B" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("Contract Terms").Range("B1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
            If ActiveMonthName = "January" Then
                Master.Sheets("Price File").Range("J12:J" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "February" Then
                Master.Sheets("Price File").Range("M12:M" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "March" Then
                Master.Sheets("Price File").Range("P12:P" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "April" Then
                Master.Sheets("Price File").Range("S12:S" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "May" Then
                Master.Sheets("Price File").Range("V12:V" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "June" Then
                Master.Sheets("Price File").Range("Y12:Y" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "July" Then
                Master.Sheets("Price File").Range("AB12:AB" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "August" Then
                Master.Sheets("Price File").Range("AE12:AE" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "September" Then
                Master.Sheets("Price File").Range("AH12:AH" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "October" Then
                Master.Sheets("Price File").Range("AK12:AK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "November" Then
                Master.Sheets("Price File").Range("AN12:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "December" Then
                Master.Sheets("Price File").Range("G12:G" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("Contract Terms").Range("C1").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
        End If
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=59, Criteria1:=Array("BPO", "DCS", "STX", "STO", "DCD", "STC", "NFX", "DCX", "REG", "NEW", "CHG", "NWR", "NYA", "CRE", "XDK", "AQA", "TGM", "DIR", "PRV", "HOL"), Operator:=xlFilterValues
        Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=3, Criteria1:="<>"
        If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        
        Master.Sheets.Add(After:=Sheets("Settings")).Name = "FPE Terms"
        Master.Sheets("FPE Terms").Range("A1").Value = "Q"
        Master.Sheets("FPE Terms").Range("B1").Value = "Quote"
        Master.Sheets("FPE Terms").Range("C1").Value = Branchcode & "/" & CurrentAccountTerms
        Master.Sheets("FPE Terms").Range("D1").Value = "Customer - " & CurrentAccountTerms
        Master.Sheets("FPE Terms").Range("E1").Value = CurrentAccountTerms
        Master.Sheets("FPE Terms").Range("F1").Value = "q0"
        Master.Sheets("FPE Terms").Range("G1").Value = "39352"
        Master.Sheets("FPE Terms").Range("H1").Value = "14:45:46"
        Master.Sheets("FPE Terms").Range("I1").Value = "almir"
        
        Master.Sheets("FPE Terms").Range("A2").Value = "Item"
        Master.Sheets("FPE Terms").Range("B2").Value = "Description"
        Master.Sheets("FPE Terms").Range("C2").Value = "Discount 1"
        Master.Sheets("FPE Terms").Range("D2").Value = "Discount 2"
        Master.Sheets("FPE Terms").Range("E2").Value = "Quantity"
        Master.Sheets("FPE Terms").Range("F2").Value = "Per(Qty)"
        Master.Sheets("FPE Terms").Range("G2").Value = "Price"
        Master.Sheets("FPE Terms").Range("H2").Value = "Per(Price)"
        
        Master.Sheets("Price File").Range("A12:A" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
        Master.Sheets("FPE Terms").Range("A3").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
            If ActiveMonthName = "January" Then
                Master.Sheets("Price File").Range("J12:J" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "February" Then
                Master.Sheets("Price File").Range("M12:M" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "March" Then
                Master.Sheets("Price File").Range("P12:P" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "April" Then
                Master.Sheets("Price File").Range("S12:S" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "May" Then
                Master.Sheets("Price File").Range("V12:V" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "June" Then
                Master.Sheets("Price File").Range("Y12:Y" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "July" Then
                Master.Sheets("Price File").Range("AB12:AB" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "August" Then
                Master.Sheets("Price File").Range("AE12:AE" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "September" Then
                Master.Sheets("Price File").Range("AH12:AH" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "October" Then
                Master.Sheets("Price File").Range("AK12:AK" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "November" Then
                Master.Sheets("Price File").Range("AN12:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            If ActiveMonthName = "December" Then
                Master.Sheets("Price File").Range("G12:G" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
                Master.Sheets("FPE Terms").Range("G3").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            Else
        End If
            Else
        Master.Sheets("Price File").AutoFilter.ShowAllData
        End If
        Master.Sheets("Price File").AutoFilter.ShowAllData
        LastRowTerms = Master.Sheets("FPE Terms").Range("A" & Rows.Count).End(xlUp).Row
        Master.Sheets("FPE Terms").Range("E3:E" & LastRowTerms).Value = "1"
        If Master.Sheets("Settings").Range("B2").Value = "BOTH" Then
            Master.Sheets("FPE Terms").Copy
            ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Terms\" & FutureMonthNameDir & "\" & Branchcode & " - " & CurrentAccount & " " & AccountFullName & " " & FutureMonthName & " - FPE Terms.csv", FileFormat:=xlCSV
            ActiveWorkbook.Close True
            Master.Sheets("Contract Terms").Copy
            ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Terms\" & FutureMonthNameDir & "\" & ContractName & " - " & AccountFullName & " " & FutureMonthName & " - Contract Terms.xlsx"
            ActiveWorkbook.Close True
        End If
    End If
    
            If ActiveMonthName = "January" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=12, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "February" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=15, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "March" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=18, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "April" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=21, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "May" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=24, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "June" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=27, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "July" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=30, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "August" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=33, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "September" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=36, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "October" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=39, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "November" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=42, Criteria1:="<>0%"
            Else
        End If
        If ActiveMonthName = "December" Then
                Master.Sheets("Price File").Range("A12:BV" & LastRowMaster).AutoFilter Field:=9, Criteria1:="<>0%"
            Else
        End If
        
        Master.Sheets.Add(After:=Sheets("Settings")).Name = "Changes"
        Master.Sheets("Price File").Range("A10:AP" & LastRowMaster).Copy
        Master.Sheets("Changes").Range("A1").PasteSpecial Paste:=xlPasteValues
        Master.Sheets("Changes").Range("A1").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        LastRowChanges = Master.Sheets("Changes").Range("A" & Rows.Count).End(xlUp).Row
        Master.Sheets("Changes").Columns(6).Delete
        Master.Sheets("Changes").Columns(7).Delete
        Master.Sheets("Changes").Columns(9).Delete
        Master.Sheets("Changes").Columns(11).Delete
        Master.Sheets("Changes").Columns(13).Delete
        Master.Sheets("Changes").Columns(15).Delete
        Master.Sheets("Changes").Columns(17).Delete
        Master.Sheets("Changes").Columns(19).Delete
        Master.Sheets("Changes").Columns(21).Delete
        Master.Sheets("Changes").Columns(23).Delete
        Master.Sheets("Changes").Columns(25).Delete
        Master.Sheets("Changes").Columns(27).Delete
        Master.Sheets("Changes").Columns(29).Delete
        If ActiveMonthName = "January" Then
                Master.Sheets("Changes").Range("J:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "February" Then
                Master.Sheets("Changes").Range("F:G,L:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "March" Then
                Master.Sheets("Changes").Range("F:I,N:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "April" Then
                Master.Sheets("Changes").Range("F:K,P:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "May" Then
                Master.Sheets("Changes").Range("F:M,R:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "June" Then
                Master.Sheets("Changes").Range("F:O,T:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "July" Then
                Master.Sheets("Changes").Range("F:Q,V:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "August" Then
                Master.Sheets("Changes").Range("F:S,X:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "September" Then
                Master.Sheets("Changes").Range("F:U,Z:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "October" Then
                Master.Sheets("Changes").Range("F:W,AB:AC").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "November" Then
                Master.Sheets("Changes").Range("F:Y").EntireColumn.Delete
            Else
        End If
        If ActiveMonthName = "December" Then
                Master.Sheets("Changes").Range("H:AA").EntireColumn.Delete
            Else
        End If
        Master.Sheets("Changes").Range("A2:I" & LastRowChanges).AutoFilter
        Master.Sheets("Changes").Columns("A:I").AutoFit
        ActiveWindow.DisplayGridlines = False
        Master.Sheets("Changes").Copy
        ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Changes\" & FutureMonthNameDir & "\" & ContractName & " - " & AccountFullName & " " & FutureMonthName & " - Changes.xlsx", FileFormat:=51
    ActiveWorkbook.Close True
    Master.Close False
End Sub
Sub Export_CustomerVersion()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.Zoom = 85
    ActiveWindow.DisplayGridlines = False
    
    Dim Master As ThisWorkbook
    Dim Additions As Workbook
    Dim Estimator As String
    Dim CurrentAccount As String
    Dim ActiveMonthName As String
    Dim PricePointUpload As Workbook
    Dim answer As Integer
    
    Set Master = ThisWorkbook
    
    Master.Sheets("Price File").Columns.EntireColumn.Hidden = False
    Master.Sheets("Price File").Rows.EntireRow.Hidden = False
    
    Estimator = Master.Sheets("Settings").Range("B1").Value
    AccountName = Master.Sheets("Price File").Range("B1").Value
    CurrentAccount = Master.Sheets("Price File").Range("B3").Value
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    AccountFullName = Master.Sheets("Price File").Range("B1").Value
    Master.Activate
    
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=48, Criteria1:="*rem*"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets.Add.Name = "REMOVALS"
        Master.Sheets("Price File").Range("A11:E" & LastRowMaster).Copy
        Master.Sheets("REMOVALS").Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Price File").Range("AR11:AR" & LastRowMaster).Copy
        Master.Sheets("REMOVALS").Range("F1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("REMOVALS").Columns.AutoFit
        Master.Sheets("REMOVALS").Rows.AutoFit
        Master.Sheets("Price File").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
        If Master.Sheets("Price File").AutoFilterMode = True Then Master.Sheets("Price File").AutoFilterMode = False
        Master.Sheets("REMOVALS").Copy
        ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Terms\" & FutureMonthNameDir & "\" & CurrentAccount & " " & AccountFullName & " " & FutureMonthName & " - Terms Removal.csv", FileFormat:=xlCSV
        ActiveWorkbook.Close True
        Master.Activate
        Master.Sheets("REMOVALS").Delete
        Master.Save
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row

    Master.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Customer Version"
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    Master.Sheets("Price File").Columns("A:AQ").Copy
    Master.Sheets("Customer Version").Range("A1").PasteSpecial Paste:=xlPasteValues
    Master.Sheets("Customer Version").Range("A1").PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    ActiveWindow.DisplayGridlines = False
    LastRowCustomer = Master.Sheets("Customer Version").Range("A" & Rows.Count).End(xlUp).Row
    Master.Sheets("Customer Version").Range("F:F,H:H,K:K,N:N,Q:Q,T:T,W:W,Z:Z,AC:AC,AF:AF,AI:AI,AL:AL,AO:AO").EntireColumn.Delete
    Master.Sheets("Customer Version").Range("A11:AD" & LastRowCustomer).AutoFilter
    Master.Sheets("Customer Version").Columns.AutoFit
    Master.Sheets("Customer Version").Copy
    ActiveSheet.Range("D3").Select
    ActiveSheet.Pictures.Insert("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Logo.jpg").Select
    ActiveWorkbook.SaveAs "\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Customer Price Files\" & FutureMonthNameDir & "\" & AccountName & " - " & CurrentAccount & " Customer Version.xlsx"
    ActiveWorkbook.Close False
    Master.Close False
End Sub
Sub Estimator_Helper()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook
    
    Master.Sheets("Price File").Columns.EntireColumn.Hidden = False
    Master.Sheets("Price File").Rows.EntireRow.Hidden = False
    
    LastRowMaster = Master.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    ActiveMonthNameDir = Format(DateSerial(Year(Date), Month(Date), 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthNameDir = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mm") & " " & Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    ActiveMonthName = Format(DateSerial(Year(Date), Month(Date), 1), "mmmm")
    FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
    LastMonthName = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mmmm")
    
    Dim Spin As Workbook

    Set Spin = Workbooks.Open("\\wukrls00fp001\RLS_Data\Departments\NACET\Price File Maintenance\Spin File\" & FutureMonthNameDir & "\SPIN.xlsx", ReadOnly:=True)
    LastRowSpin = Spin.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterNextMonth, Operator:=xlFilterDynamic
        If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Next Month"
                Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
                Spin.Sheets("Next Month").Range("A1").PasteSpecial Paste:=xlPasteValues
                Spin.Sheets("Next Month").Range("A1").PasteSpecial Paste:=xlFormats
                Application.CutCopyMode = False
                Spin.Sheets(1).AutoFilter.ShowAllData
            Else
        Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterThisMonth, Operator:=xlFilterDynamic
    If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "This Month"
            Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
            Spin.Sheets("This Month").Range("A1").PasteSpecial Paste:=xlPasteValues
            Spin.Sheets("This Month").Range("A1").PasteSpecial Paste:=xlFormats
            Application.CutCopyMode = False
            Spin.Sheets(1).AutoFilter.ShowAllData
        Else
    Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Spin.Sheets(1).Range("A1:J" & LastRowSpin).AutoFilter Field:=7, Criteria1:=xlFilterLastMonth, Operator:=xlFilterDynamic
    If Spin.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Spin.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Last Month"
            Spin.Sheets(1).Range("A1:J" & LastRowSpin).SpecialCells(xlCellTypeVisible).Copy
            Spin.Sheets("Last Month").Range("A1").PasteSpecial Paste:=xlPasteValues
            Spin.Sheets("Last Month").Range("A1").PasteSpecial Paste:=xlFormats
            Application.CutCopyMode = False
            Spin.Sheets(1).AutoFilter.ShowAllData
        Else
    Spin.Sheets(1).AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:=ActiveMonthName & " Increase"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]This Month'!$C:$J,8,FALSE)"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:=LastMonthName & " Increase"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]Last Month'!$C:$J,8,FALSE)"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter Field:=65, Criteria1:=FutureMonthName & " Increase"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@BJ:BJ,'[SPIN.xlsx]Next Month'!$C:$J,8,FALSE)"
            Master.Sheets("Price File").AutoFilter.ShowAllData
        Else
    Master.Sheets("Price File").AutoFilter.ShowAllData
    End If
    Master.Sheets("Price File").AutoFilter.ShowAllData
    Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).Copy
    Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).NumberFormat = "0.00%"
    Master.Sheets("Price File").AutoFilterMode = False
    Spin.Close False
    Master.Sheets("Price File").Range("A11:BY" & LastRowMaster).AutoFilter Field:=77, Criteria1:="#N/A"
    If Master.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
        Master.Sheets("Price File").Range("BY12:BY" & LastRowMaster).SpecialCells(xlCellTypeVisible).ClearContents
    End If
    Master.Sheets("Price File").AutoFilterMode = False
    Master.Sheets("Price File").Range("BY11").Value = "Max Spin"
    Master.Sheets("Price File").Range("BZ11").Value = "COND1"
    Master.Sheets("Price File").Range("CA11").Value = "COND2"
    Master.Sheets("Price File").Range("CB11").Value = "COND3"
    Master.Sheets("Price File").Range("CC11").Value = "COND4"
    Master.Sheets("Price File").Range("CD11").Value = "COND5"
    Master.Sheets("Price File").Range("CE11").Value = "COND6"
    Master.Sheets("Price File").Range("CF11").Value = "COND6"
    Master.Sheets("Price File").Range("A11:CF" & LastRowMaster).AutoFilter
    
    If FutureMonthName = "January" Then
        ONEMONTHAGO = "AP12"
        TWOMONTHAGO = "AM12"
    End If
    If FutureMonthName = "February" Then
        ONEMONTHAGO = "I12"
        TWOMONTHAGO = "AP12"
    End If
    If FutureMonthName = "March" Then
        ONEMONTHAGO = "I12"
        TWOMONTHAGO = "L12"
    End If
    If FutureMonthName = "April" Then
        ONEMONTHAGO = "L12"
        TWOMONTHAGO = "O12"
    End If
    If FutureMonthName = "May" Then
        ONEMONTHAGO = "O12"
        TWOMONTHAGO = "R12"
    End If
    If FutureMonthName = "June" Then
        ONEMONTHAGO = "R12"
        TWOMONTHAGO = "U12"
    End If
    If FutureMonthName = "July" Then
        ONEMONTHAGO = "U12"
        TWOMONTHAGO = "X12"
    End If
    If FutureMonthName = "August" Then
        ONEMONTHAGO = "X12"
        TWOMONTHAGO = "AA12"
    End If
    If FutureMonthName = "September" Then
        ONEMONTHAGO = "AA12"
        TWOMONTHAGO = "AD12"
    End If
    If FutureMonthName = "October" Then
        ONEMONTHAGO = "AD12"
        TWOMONTHAGO = "AG12"
    End If
    If FutureMonthName = "November" Then
        ONEMONTHAGO = "AG12"
        TWOMONTHAGO = "AJ12"
    End If
    If FutureMonthName = "December" Then
        ONEMONTHAGO = "AJ12"
        TWOMONTHAGO = "AM12"
    End If
    
    
    If Master.Sheets("Settings").Range("B7").Value = "SLP" Then
        Master.Sheets("Price File").Range("BZ12:BZ" & LastRowMaster).Formula = "=IF(@BE:BE>0,""[Increase]"",IF(@BE:BE<0,""[Decrease]"",IF(@BE:BE=0,""[No Change]"",""[ERROR]"")))"
        Master.Sheets("Price File").Range("CA12:CA" & LastRowMaster).Formula = "=IF(BZ12=""[Decrease]"",""[Manual Review]"",IF(BM12=""Not Inline Increase"",""[Not Inline Spin]"",IF(BZ12=""[No Change]"",""[No Change]"",IF(AND(OR(ISBLANK(BY12),BY12=""""),BB12<>0),""[Manual Review]"",IF((BB12-BY12)<1%,""[Inline]"",""[Not Inline]"")))))"
        Master.Sheets("Price File").Range("CB12:CB" & LastRowMaster).Formula = "=IF(AND(NOT(ISNUMBER(AZ12)),ISNUMBER(BA12)),""[Added Support]"",IF(AND(ISNUMBER(AZ12),NOT(ISNUMBER(BA12))),""[Expired Support]"",IF(AND(ISNUMBER(BF12),BF12>0),""[Increased Support]"",IF(AND(ISNUMBER(BF12),BF12<0),""[Decreased Support]"",IF(AND(ISNUMBER(AZ12),ISNUMBER(BA12),AZ12=BA12,BB12>0),""[Offset Support Required]"",IF(OR(AZ12=BA12,BB12=""No Data""),""[No Change]""))))))"
        Master.Sheets("Price File").Range("CC12:CC" & LastRowMaster).Formula = "=IF(AND(ISNUMBER(BU12),ISNUMBER(BT12),BU12=BT12,AZ12=BA12,BP12<>BQ12,CB12<>""[Offset Support Required]""),""[Internal Change]"",""[No Change]"")"
        Master.Sheets("Price File").Range("CD12:CD" & LastRowMaster).Formula = "=IFERROR(IF(AND(ISNUMBER(BY12),ISNUMBER(BB12),AQ12<>""No Usage"",AT12>10,(BY12-BB12)>1%,BB12>0,CB12<>""[Offset Support Required]""),""[Can Boost]"",""[No Change]""),""[No Change]"")"
        Master.Sheets("Price File").Range("CE12:CE" & LastRowMaster).Formula = "=IF(AND(BD12<>""No Data"",BD12<-0.01,AZ12=BA12),""[Margin Loss]"",""[No Change]"")"
        Master.Sheets("Price File").Range("CF12:CF" & LastRowMaster).Formula = "=IF(AND(OR(" & ONEMONTHAGO & "<>0," & TWOMONTHAGO & "<>0),BB12<>0,BB12<>""No Data""),""[Recently Changed]"",""[No Change]"")"
        Master.Sheets("Price File").Range("BW12:BW" & LastRowMaster).Formula = "=SUBSTITUTE(BZ12&CA12&CB12&CC12&CD12&CE12&CF12,""[No Change]"","""")"
        Master.Sheets("Price File").Range("BW12:BW" & LastRowMaster).Copy
        Master.Sheets("Price File").Range("BW12:BW" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    If Master.Sheets("Settings").Range("B7").Value = "MARGIN" Then
        Master.Sheets("Price File").Range("BZ12:BZ" & LastRowMaster).Formula = "=IF(BB12=""No Data"",""[No Change]"",IF(@BB:BB>0,""[SLP Increase Next month]"",IF(@BB:BB<0,""[No Change]"",IF(@BB:BB=0,""[No Change]"",""[ERROR]""))))"
        Master.Sheets("Price File").Range("CB12:CB" & LastRowMaster).Formula = "=IF(BZ12=""[SLP Increase Next Month]"",""[No Change]"",IF(AND(NOT(ISNUMBER(AZ12)),ISNUMBER(BA12)),""[Added Support]"",IF(AND(ISNUMBER(AZ12),NOT(ISNUMBER(BA12))),""[Expired Support]"",IF(AND(ISNUMBER(BF12),BF12>0),""[Increased Support]"",IF(AND(ISNUMBER(BF12),BF12<0),""[Decreased Support]"",IF(AND(ISNUMBER(AZ12),ISNUMBER(BA12),AZ12=BA12,BB12>0),""[Offset Support Required]"",IF(OR(AZ12=BA12,BB12=""No Data""),""[No Change]"")))))))"
        Master.Sheets("Price File").Range("CC12:CC" & LastRowMaster).Formula = "=IF(AND(ISNUMBER(BU12),ISNUMBER(BT12),BU12=BT12,AZ12=BA12,BP12<>BQ12,CB12<>""[Offset Support Required]""),""[Internal Change]"",""[No Change]"")"
        Master.Sheets("Price File").Range("BW12:BW" & LastRowMaster).Formula = "=SUBSTITUTE(BZ12&CB12&CC12,""[No Change]"","""")"
        Master.Sheets("Price File").Range("BW12:BW" & LastRowMaster).Copy
        Master.Sheets("Price File").Range("BW12:BW" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    Master.Sheets("Price File").Range("BY11:CF" & LastRowMaster).ClearContents
    Master.Sheets("Price File").AutoFilterMode = False
    Master.Sheets("Price File").Range("A11:BX" & LastRowMaster).AutoFilter
    Master.Sheets("Price File").Columns.AutoFit
    Master.Sheets("Price File").Rows.AutoFit
        If ActiveMonthName = "January" Then
        Master.Sheets("Price File").Columns("M:AP").Hidden = True
        Else
    End If
        If ActiveMonthName = "February" Then
        Master.Sheets("Price File").Columns("P:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:I").Hidden = True
        Else
    End If
        If ActiveMonthName = "March" Then
        Master.Sheets("Price File").Columns("S:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:L").Hidden = True
        Else
    End If
        If ActiveMonthName = "April" Then
        Master.Sheets("Price File").Columns("V:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:O").Hidden = True
        Else
    End If
        If ActiveMonthName = "May" Then
        Master.Sheets("Price File").Columns("Y:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:R").Hidden = True
        Else
    End If
        If ActiveMonthName = "June" Then
        Master.Sheets("Price File").Columns("AB:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:U").Hidden = True
        Else
    End If
        If ActiveMonthName = "July" Then
        Master.Sheets("Price File").Columns("AE:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:X").Hidden = True
        Else
    End If
        If ActiveMonthName = "August" Then
        Master.Sheets("Price File").Columns("AH:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:AA").Hidden = True
        Else
    End If
        If ActiveMonthName = "September" Then
        Master.Sheets("Price File").Columns("AK:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:AD").Hidden = True
        Else
    End If
        If ActiveMonthName = "October" Then
        Master.Sheets("Price File").Columns("AN:AP").Hidden = True
        Master.Sheets("Price File").Columns("G:AD").Hidden = True
        Else
    End If
        If ActiveMonthName = "November" Then
        Master.Sheets("Price File").Columns("G:AJ").Hidden = True
        Else
    End If
        If ActiveMonthName = "December" Then
        Master.Sheets("Price File").Columns("J:AM").Hidden = True
        Else
    End If
    Master.Save
End Sub
Function WUKCODE(file_name As String) As String
Dim regEx As Object
Set regEx = CreateObject("vbscript.regexp")

regEx.Pattern = "[A-Za-z]{2}[0-9]{4}"
If regEx.Test(file_name) Then
    WUKCODE = regEx.Execute(file_name)(0)
Else
    regEx.Pattern = "[A-Za-z]{1}[0-9]{5}"
    If regEx.Test(file_name) Then
        WUKCODE = regEx.Execute(file_name)(0)
    Else
        regEx.Pattern = "[0-9]{6}"
        If regEx.Test(file_name) Then
            WUKCODE = regEx.Execute(file_name)(0)
        Else
        WUKCODE = "Unknown Code"
        End If
    End If
End If
End Function
