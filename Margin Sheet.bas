Sub Margin_sheet()
' Last Update: 24.09.2020 by Andrei Polin
SHEETNAME = "Finance Margins"
Pricepoint = "Price Point"
user = "Andrei Polin"

Application.DisplayAlerts = False
Application.ScreenUpdating = False
    Sheets(1).Name = Pricepoint
    On Error Resume Next
    Worksheets(SHEETNAME).Delete
    On Error GoTo 0
    Worksheets.Add(After:=Sheets(Sheets.Count)).Name = SHEETNAME
    Sheets(SHEETNAME).Range("A10").Value = "Supplier Part No."
    Sheets(SHEETNAME).Range("B10").Value = "Wolseley Code"
    Sheets(SHEETNAME).Range("C10").Value = "Product Description"
    Sheets(SHEETNAME).Range("D10").Value = "Terms Price (£)"
    Sheets(SHEETNAME).Range("E10").Value = "Quantity"
    Sheets(SHEETNAME).Range("F10").Value = "Margin (%)"
    Sheets(SHEETNAME).Range("G10").Value = "Support (£)"
    Sheets(SHEETNAME).Range("H10").Value = "WUK Ref."
    Sheets(SHEETNAME).Range("I10").Value = "Nett Cost (£)"
    Sheets(SHEETNAME).Range("J10").Value = "Nett Cost NRA (£)"
    Sheets(SHEETNAME).Range("K10").Value = "Total Sell (£)"
    Sheets(SHEETNAME).Range("L10").Value = "Total Cost (£)"
    Sheets(SHEETNAME).Range("M10").Value = "Total Profit (£)"
    Sheets(SHEETNAME).Range("N10").Value = "Set Margin (%)"
    Sheets(SHEETNAME).Range("O10").Value = "New Sell (£)"
    Sheets(SHEETNAME).Range("P10").Value = "New Discount (%)"
    Sheets(SHEETNAME).Range("Q10").Value = "Set Discount (%)"
    Sheets(SHEETNAME).Range("R10").Value = "New Sell (£)"
    Sheets(SHEETNAME).Range("S10").Value = "New Margin (%)"
    Sheets(SHEETNAME).Range("T10").Value = "Wolseley SPG"
    Sheets(SHEETNAME).Range("U10").Value = "Supplier Name"
    Sheets(SHEETNAME).Range("V10").Value = "Supplier HoS"
    Sheets(SHEETNAME).Range("W10").Value = "Current Invoice (£)"
    Sheets(SHEETNAME).Range("X10").Value = "Future Invoice (£)"
    Sheets(SHEETNAME).Range("Y10").Value = "Invoice Change (%)"
    Sheets(SHEETNAME).Range("Z10").Value = "Current Trade (£)"
    Sheets(SHEETNAME).Range("AA10").Value = "Future Trade (£)"
    Sheets(SHEETNAME).Range("AB10").Value = "Trade Change (%)"
    Sheets(SHEETNAME).Range("AC10").Value = "Current Branch (£)"
    Sheets(SHEETNAME).Range("AD10").Value = "Future Branch (£)"
    Sheets(SHEETNAME).Range("AE10").Value = "Branch Change (%)"
    Sheets(SHEETNAME).Range("AF10").Value = "Current SLP (£)"
    Sheets(SHEETNAME).Range("AG10").Value = "Future SLP (£)"
    Sheets(SHEETNAME).Range("AH10").Value = "SLP Change (%)"
    Sheets(SHEETNAME).Range("AI10").Value = "Current Real Cost (£)"
    Sheets(SHEETNAME).Range("AJ10").Value = "Future Real Cost (£)"
    Sheets(SHEETNAME).Range("AK10").Value = "Real Change (%)"
    Sheets(SHEETNAME).Range("AL10").Value = "Future Date"
    Sheets(SHEETNAME).Range("AM10").Value = "Rebate Impacted"
    Sheets(SHEETNAME).Range("AN10").Value = "Product Lifecycle"
    Sheets(SHEETNAME).Range("AO10").Value = "Product Narrative"
    Sheets(SHEETNAME).Range("AP10").Value = "Last Sale Date"
    Sheets(SHEETNAME).Range("AQ10").Value = "Last Sale Price (£)"
    
    LastRow = Sheets(Pricepoint).Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(Pricepoint).Range("E15:E" & LastRow).Copy
    Sheets(SHEETNAME).Range("A11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("H15:H" & LastRow).Copy
    Sheets(SHEETNAME).Range("B11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("I15:I" & LastRow).Copy
    Sheets(SHEETNAME).Range("C11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Sheets(Pricepoint).Range("AB15:AB" & LastRow).Copy
    Sheets(SHEETNAME).Range("D11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("Z15:Z" & LastRow).Copy
    Sheets(SHEETNAME).Range("E11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("AN15:AN" & LastRow).Copy
    Sheets(SHEETNAME).Range("G11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("AJ15:AJ" & LastRow).Copy
    Sheets(SHEETNAME).Range("H11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("S15:S" & LastRow).Copy
    Sheets(SHEETNAME).Range("I11").PasteSpecial Paste:=xlPasteValues
    Sheets(SHEETNAME).Range("J11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("G15:G" & LastRow).Copy
    Sheets(SHEETNAME).Range("T11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Sheets(Pricepoint).Range("D15:D" & LastRow).Copy
    Sheets(SHEETNAME).Range("U11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("Q15:Q" & LastRow).Copy
    Sheets(SHEETNAME).Range("W11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("W15:W" & LastRow).Copy
    Sheets(SHEETNAME).Range("X11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("O15:O" & LastRow).Copy
    Sheets(SHEETNAME).Range("Z11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("U15:U" & LastRow).Copy
    Sheets(SHEETNAME).Range("AA11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("P15:P" & LastRow).Copy
    Sheets(SHEETNAME).Range("AC11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Sheets(Pricepoint).Range("V15:V" & LastRow).Copy
    Sheets(SHEETNAME).Range("AD11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("R15:R" & LastRow).Copy
    Sheets(SHEETNAME).Range("AF11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("X15:X" & LastRow).Copy
    Sheets(SHEETNAME).Range("AG11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("S15:S" & LastRow).Copy
    Sheets(SHEETNAME).Range("AI11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("Y15:Y" & LastRow).Copy
    Sheets(SHEETNAME).Range("AJ11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("T15:T" & LastRow).Copy
    Sheets(SHEETNAME).Range("AL11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("F15:F" & LastRow).Copy
    Sheets(SHEETNAME).Range("AM11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("J15:J" & LastRow).Copy
    Sheets(SHEETNAME).Range("AN11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("M15:M" & LastRow).Copy
    Sheets(SHEETNAME).Range("AO11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("L15:L" & LastRow).Copy
    Sheets(SHEETNAME).Range("AQ11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets(Pricepoint).Range("K15:K" & LastRow).Copy
    Sheets(SHEETNAME).Range("AP11").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    
    LastRowFM = Sheets(SHEETNAME).Columns(2).Cells(Rows.Count, 1).End(xlUp).Row
    Sheets(SHEETNAME).Range("F11:F" & LastRowFM).FormulaR1C1 = "=(RC[-2]-RC[4])/RC[-2]"
    Sheets(SHEETNAME).Range("K11:K" & LastRowFM).FormulaR1C1 = "=RC[-7]*RC[-6]"
    Sheets(SHEETNAME).Range("L11:L" & LastRowFM).FormulaR1C1 = "=RC[-2]*RC[-7]"
    Sheets(SHEETNAME).Range("M11:M" & LastRowFM).FormulaR1C1 = "=RC[-2]-RC[-1]"
    Sheets(SHEETNAME).Range("O11:O" & LastRowFM).FormulaR1C1 = "=RC[-5]/(1-RC[-1])"
    Sheets(SHEETNAME).Range("P11:P" & LastRowFM).FormulaR1C1 = "=(RC[-1]/RC[10])-1"
    Sheets(SHEETNAME).Range("R11:R" & LastRowFM).FormulaR1C1 = "=RC[8]*(1-RC[-1])"
    Sheets(SHEETNAME).Range("S11:S" & LastRowFM).FormulaR1C1 = "=(RC[-1]-RC[-9])/RC[-1]"
    Sheets(SHEETNAME).Range("Y11:Y" & LastRowFM).FormulaR1C1 = "=IFERROR((RC[-1]/RC[-2])-1,"""")"
    Sheets(SHEETNAME).Range("AB11:AB" & LastRowFM).FormulaR1C1 = "=IFERROR((RC[-1]/RC[-2])-1,"""")"
    Sheets(SHEETNAME).Range("AE11:AE" & LastRowFM).FormulaR1C1 = "=IFERROR((RC[-1]/RC[-2])-1,"""")"
    Sheets(SHEETNAME).Range("AH11:AH" & LastRowFM).FormulaR1C1 = "=IFERROR((RC[-1]/RC[-2])-1,"""")"
    Sheets(SHEETNAME).Range("AK11:AK" & LastRowFM).FormulaR1C1 = "=IFERROR((RC[-1]/RC[-2])-1,"""")"
    
    Sheets(SHEETNAME).Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Font.Underline = True
    Range(Selection, Selection.End(xlToRight)).Font.Bold = True
    Sheets(SHEETNAME).Cells.Font.Name = "Arial"
    Sheets(SHEETNAME).Cells.Font.Size = 10

    Sheets(SHEETNAME).Range("D11:D" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("G11:G" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("I11:I" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("J11:J" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("K11:K" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("L11:L" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("M11:M" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("O11:O" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("R11:R" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("W11:W" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("X11:X" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("Z11:Z" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AQ11:AQ" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AA11:AA" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AC11:AC" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AD11:AD" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AF11:AF" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AG11:AG" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AI11:AI" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("AJ11:AJ" & LastRowFM).NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("F4").NumberFormat = "£#,##0.00"
    Sheets(SHEETNAME).Range("F11:F" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("P11:P" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("Q11:Q" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("S11:S" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("Y11:Y" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("AB11:AB" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("AE11:AE" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("AH11:AH" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("AK11:AK" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("N11:N" & LastRowFM).NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("F3").NumberFormat = "0.00%"
    Sheets(SHEETNAME).Range("B11:B" & LastRowFM).TextToColumns Destination:=Range("B11"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Sheets(SHEETNAME).Range("A11:A" & LastRowFM).TextToColumns Destination:=Range("A11"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets(SHEETNAME).Range("F9:J9").Value = "Finance Margins"
    Sheets(SHEETNAME).Range("F9:J9").Merge
    Sheets(SHEETNAME).Range("F9:J9").Font.Bold = True
    Sheets(SHEETNAME).Range("F9:J9").HorizontalAlignment = xlCenter
    Sheets(SHEETNAME).Range("AL11:AL" & LastRowFM).NumberFormat = "dd/mm/yyyy"
    Sheets(SHEETNAME).Range("AP11:AP" & LastRowFM).NumberFormat = "dd/mm/yyyy"
    
    Sheets(SHEETNAME).Range("K9:M9").Value = "Basket Totals"
    Sheets(SHEETNAME).Range("K9:M9").Merge
    Sheets(SHEETNAME).Range("K9:M9").Font.Bold = True
    Sheets(SHEETNAME).Range("K9:M9").HorizontalAlignment = xlCenter
    
    Sheets(SHEETNAME).Range("N9:P9").Value = "Cost Plus Calculator"
    Sheets(SHEETNAME).Range("N9:P9").Merge
    Sheets(SHEETNAME).Range("N9:P9").Font.Bold = True
    Sheets(SHEETNAME).Range("N9:P9").HorizontalAlignment = xlCenter
    
    Sheets(SHEETNAME).Range("Q9:S9").Value = "Discount Off Calculator"
    Sheets(SHEETNAME).Range("Q9:S9").Merge
    Sheets(SHEETNAME).Range("Q9:S9").Font.Bold = True
    Sheets(SHEETNAME).Range("Q9:S9").HorizontalAlignment = xlCenter
    
    Sheets(SHEETNAME).Range("T9:V9").Value = "Product Details"
    Sheets(SHEETNAME).Range("T9:V9").Merge
    Sheets(SHEETNAME).Range("T9:V9").Font.Bold = True
    Sheets(SHEETNAME).Range("T9:V9").HorizontalAlignment = xlCenter
    
    Sheets(SHEETNAME).Range("W9:AQ9").Value = "Pricing Details"
    Sheets(SHEETNAME).Range("W9:AQ9").Merge
    Sheets(SHEETNAME).Range("W9:AQ9").Font.Bold = True
    Sheets(SHEETNAME).Range("W9:AQ9").HorizontalAlignment = xlCenter
    
    Sheets(SHEETNAME).Range("A9:E9").Value = "Finance Margin Sheet - Generated Via Price Point"
    Sheets(SHEETNAME).Range("A9:E9").Merge
    Sheets(SHEETNAME).Range("A9:E9").Font.Bold = True
    Sheets(SHEETNAME).Range("A9:E9").HorizontalAlignment = xlCenter
    
    Sheets(SHEETNAME).Range("A10:AQ10").Interior.Color = RGB(150, 194, 230)
    Sheets(SHEETNAME).Range("A9:E9").Interior.Color = RGB(237, 237, 237)
    Sheets(SHEETNAME).Range("A11:E" & LastRowFM).Interior.Color = RGB(237, 237, 237)
    Sheets(SHEETNAME).Range("W9:AK9").Interior.Color = RGB(237, 237, 237)
    Sheets(SHEETNAME).Range("W11:AK" & LastRowFM).Interior.Color = RGB(237, 237, 237)
    Sheets(SHEETNAME).Range("F9:J9").Interior.Color = RGB(198, 224, 180)
    Sheets(SHEETNAME).Range("F11:J" & LastRowFM).Interior.Color = RGB(198, 224, 180)
    Sheets(SHEETNAME).Range("K9:M9").Interior.Color = RGB(146, 208, 80)
    Sheets(SHEETNAME).Range("K11:M" & LastRowFM).Interior.Color = RGB(146, 208, 80)
    Sheets(SHEETNAME).Range("N9:P9").Interior.Color = RGB(255, 204, 204)
    Sheets(SHEETNAME).Range("N11:P" & LastRowFM).Interior.Color = RGB(255, 204, 204)
    Sheets(SHEETNAME).Range("Q9:S9").Interior.Color = RGB(248, 203, 173)
    Sheets(SHEETNAME).Range("Q11:S" & LastRowFM).Interior.Color = RGB(248, 203, 173)
    Sheets(SHEETNAME).Range("T9:V9").Interior.Color = RGB(242, 242, 242)
    Sheets(SHEETNAME).Range("T11:V" & LastRowFM).Interior.Color = RGB(242, 242, 242)
    Sheets(SHEETNAME).Range("AL9:AQ9").Interior.Color = RGB(242, 242, 242)
    Sheets(SHEETNAME).Range("AL11:AQ" & LastRowFM).Interior.Color = RGB(242, 242, 242)
    Sheets(SHEETNAME).Range("A9:AQ" & LastRowFM).Borders.LineStyle = xlContinuous
    Sheets(SHEETNAME).Range("A9:AQ" & LastRowFM).Borders.Weight = xlThin
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).HorizontalAlignment = xlCenter
    Sheets(SHEETNAME).Range("F11:F" & LastRowFM).Font.Bold = True
    Sheets(SHEETNAME).Range("P11:P" & LastRowFM).Font.Bold = True
    Sheets(SHEETNAME).Range("S11:S" & LastRowFM).Font.Bold = True
    
    Sheets(SHEETNAME).Range("A1").Value = "Account ID:"
    Sheets(SHEETNAME).Range("A2").Value = "Account Name:"
    Sheets(SHEETNAME).Range("A3").Value = "Generator:"
    Sheets(SHEETNAME).Range("A4").Value = "Quote ID:"
    Sheets(SHEETNAME).Range("A5").Value = "Date Generated:"
    Sheets(SHEETNAME).Range("B1").Value = Sheets(Pricepoint).Range("B2")
    Sheets(SHEETNAME).Range("B2").Value = Sheets(Pricepoint).Range("B3")
    Sheets(SHEETNAME).Range("B3").Value = Sheets(Pricepoint).Range("B11")
    Sheets(SHEETNAME).Range("B4").Value = Sheets(Pricepoint).Range("B4")
    Sheets(SHEETNAME).Range("B5").Value = Now()
    Sheets(SHEETNAME).Range("A1:B5").Borders.LineStyle = xlContinuous
    Sheets(SHEETNAME).Range("A1:B5").Borders.Weight = xlThin
    Sheets(SHEETNAME).Range("A1:A5").Font.Bold = True
    Sheets(SHEETNAME).Range("A1:B5").Interior.Color = RGB(150, 194, 230)
    Sheets(SHEETNAME).Range("B1:B5").HorizontalAlignment = xlLeft
    Sheets(SHEETNAME).Range("E1:F4").Interior.Color = RGB(255, 204, 153)
    Sheets(SHEETNAME).Range("E1:F4").Borders.LineStyle = xlContinuous
    Sheets(SHEETNAME).Range("E1:F4").Borders.Weight = xlThin
    
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=39, Criteria1:="N"
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("I11:I" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[26]"
            Sheets(SHEETNAME).Range("J11:J" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-1]-RC[-3]"
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=39, Criteria1:="<>N"
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("I11:I" & LastRowFM).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 0, 0)
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=24, Criteria1:="="
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("X11:X" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-1]"
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=27, Criteria1:="="
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("AA11:AA" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-1]"
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=30, Criteria1:="="
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("AD11:AD" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-1]"
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=36, Criteria1:="="
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("AJ11:AJ" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-1]"
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    
    Sheets(SHEETNAME).Range("A10:AQ" & LastRowFM).AutoFilter field:=33, Criteria1:="="
        If Sheets(SHEETNAME).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets(SHEETNAME).Range("AG11:AG" & LastRowFM).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-1]"
            Sheets(SHEETNAME).AutoFilter.ShowAllData
            Else
        Sheets(SHEETNAME).AutoFilter.ShowAllData
    End If
    
    Sheets(SHEETNAME).Range("E1").Value = "Total Sell:"
    Sheets(SHEETNAME).Range("E2").Value = "Total Cost:"
    Sheets(SHEETNAME).Range("E3").Value = "Basket Margin:"
    Sheets(SHEETNAME).Range("E4").Value = "Total Profit:"
    Sheets(SHEETNAME).Range("F1").Formula = "=SUM(K11:K" & LastRowFM & ")"
    Sheets(SHEETNAME).Range("F2").Formula = "=SUM(L11:L" & LastRowFM & ")"
    Sheets(SHEETNAME).Range("F3").Formula = "=(F1-F2)/F1"
    Sheets(SHEETNAME).Range("F4").Formula = "=F1*F3"
    ActiveWindow.DisplayGridlines = False
    
    Sheets(Pricepoint).Delete
    Sheets(SHEETNAME).Columns.AutoFit
    Dim LoginName As String
    Dim accountid As String
    Dim Accountname As String
    Dim accounttime As String
    Dim accounttime2 As String
    accountid = Sheets(SHEETNAME).Range("B1").Value
    Accountname = Sheets(SHEETNAME).Range("B2").Value
    accounttime = Replace(Sheets(SHEETNAME).Range("B4").Value, ".", "")
    accounttime2 = Replace(accounttime, "/", "")
    LoginName = Environ$("username")
    ActiveWorkbook.SaveAs Filename:="\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & user & "\Quotes\" & accountid & " - " & accounttime2 & " Margin Sheet" & ".xlsx"
    ActiveWorkbook.Close True
End Sub
