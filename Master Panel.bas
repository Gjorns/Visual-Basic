Sub UpdateData()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook
    Dim ControlPanel As Workbook
    Dim PriceFile As Workbook
    
    LastRowMasterData = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
    LastRowMasterEstimators = Master.Sheets("Estimators").Range("A" & Rows.Count).End(xlUp).Row
    
    Master.Sheets("No. Products").Columns.EntireColumn.Hidden = False
    Master.Sheets("No. Products").Rows.EntireRow.Hidden = False
    Master.Sheets("No. EOL").Columns.EntireColumn.Hidden = False
    Master.Sheets("No. EOL").Rows.EntireRow.Hidden = False
    Master.Sheets("No. No-Usage").Columns.EntireColumn.Hidden = False
    Master.Sheets("No. No-Usage").Rows.EntireRow.Hidden = False
    Master.Sheets("No. Internal").Columns.EntireColumn.Hidden = False
    Master.Sheets("No. Internal").Rows.EntireRow.Hidden = False
    Master.Sheets("No. Not-Inline-Spin").Columns.EntireColumn.Hidden = False
    Master.Sheets("No. Not-Inline-Spin").Rows.EntireRow.Hidden = False
    Master.Sheets("No. Not Pass").Columns.EntireColumn.Hidden = False
    Master.Sheets("No. Not Pass").Rows.EntireRow.Hidden = False
    Master.Sheets("Loss Value").Columns.EntireColumn.Hidden = False
    Master.Sheets("Loss Value").Rows.EntireRow.Hidden = False
    Master.Sheets("Basket Margin").Columns.EntireColumn.Hidden = False
    Master.Sheets("Basket Margin").Rows.EntireRow.Hidden = False
    Master.Sheets("Basket Profit").Columns.EntireColumn.Hidden = False
    Master.Sheets("Basket Profit").Rows.EntireRow.Hidden = False
    Master.Sheets("No. Products").AutoFilterMode = False
    Master.Sheets("No. EOL").AutoFilterMode = False
    Master.Sheets("No. No-Usage").AutoFilterMode = False
    Master.Sheets("No. Internal").AutoFilterMode = False
    Master.Sheets("No. Not-Inline-Spin").AutoFilterMode = False
    Master.Sheets("No. Not Pass").AutoFilterMode = False
    Master.Sheets("Loss Value").AutoFilterMode = False
    Master.Sheets("Basket Margin").AutoFilterMode = False
    Master.Sheets("Basket Profit").AutoFilterMode = False
    
    
    
    If LastRowMasterData > 1 Then
        Master.Sheets("DATA").Range("A2:F" & LastRowMasterData).ClearContents
        LastRowMasterData = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
    End If
    
    If LastRowMasterEstimators > 0 Then
        EstimatorsStart = 1
        EstimatorsFinish = LastRowMasterEstimators
        For I = EstimatorsStart To EstimatorsFinish
            SearchEstimator = Master.Sheets("Estimators").Range("A" & EstimatorsStart).Value
            Set ControlPanel = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & SearchEstimator & "\Control Panel.xlsm")
            LastRowControlPanel = ControlPanel.Sheets("Paths").Range("A" & Rows.Count).End(xlUp).Row
            If LastRowControlPanel > 2 Then
                ControlPanel.Sheets("Paths").Range("A3:C" & LastRowControlPanel).Copy
                Master.Sheets("DATA").Range("A" & LastRowMasterData + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                LastRowMasterDataNew = Master.Sheets("DATA").Range("C" & Rows.Count).End(xlUp).Row
                Master.Sheets("DATA").Range("E" & LastRowMasterData + 1 & ":E" & LastRowMasterDataNew).Value = SearchEstimator
                Master.Sheets("DATA").Range("D" & LastRowMasterData + 1 & ":D" & LastRowMasterDataNew).Formula = "=INDEX('[" & ControlPanel.Name & "]Control Panel'!$A:$A,MATCH(@A:A,'[" & ControlPanel.Name & "]Control Panel'!$B:$B,0))"
                Master.Sheets("DATA").Range("D" & LastRowMasterData + 1 & ":D" & LastRowMasterDataNew).Copy
                Master.Sheets("DATA").Range("D" & LastRowMasterData + 1 & ":D" & LastRowMasterDataNew).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            End If
            ControlPanel.Close False
            LastRowMasterData = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
            EstimatorsStart = EstimatorsStart + 1
        Next I
        Master.Sheets("DATA").Range("F2:F" & LastRowMasterData).Formula = "=B2& ""\""&C2&"".xlsm"""
        Master.Sheets("DATA").Range("F2:F" & LastRowMasterData).Copy
        Master.Sheets("DATA").Range("F2:F" & LastRowMasterData).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        Master.Sheets("DATA").Range("A1:G" & LastRowMasterData).AutoFilter
        Master.Sheets("DATA").Range("A1:G" & LastRowMasterData).AutoFilter field:=4, Criteria1:="#N/A"
        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("DATA").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Master.Sheets("DATA").AutoFilter.ShowAllData
            LastRowMasterData = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
        End If
        Master.Sheets("DATA").AutoFilter.ShowAllData
        LastRowMasterData = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
        
        LastRowMasterInfo = Master.Sheets("No. products").Range("A" & Rows.Count).End(xlUp).Row
        
        Master.Sheets("DATA").Range("G2:G" & LastRowMasterData).Formula = "=VLOOKUP(D2,'No. Products'!C:C,1,FALSE)"
        Master.Sheets("DATA").Range("A1:G" & LastRowMasterData).AutoFilter field:=7, Criteria1:="#N/A"
        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("DATA").Range("A2:A" & LastRowMasterData).Copy
            Master.Sheets("No. Products").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. EOL").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. No-Usage").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Internal").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Not-Inline-Spin").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Not Pass").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Loss Value").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Basket Margin").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Basket Profit").Range("B" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            Master.Sheets("DATA").Range("E2:E" & LastRowMasterData).Copy
            Master.Sheets("No. Products").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. EOL").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. No-Usage").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Internal").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Not-Inline-Spin").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Not Pass").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Loss Value").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Basket Margin").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Basket Profit").Range("A" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            Master.Sheets("DATA").Range("D2:D" & LastRowMasterData).Copy
            Master.Sheets("No. Products").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. EOL").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. No-Usage").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Internal").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Not-Inline-Spin").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("No. Not Pass").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Loss Value").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Basket Margin").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Master.Sheets("Basket Profit").Range("C" & LastRowMasterInfo + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End If
        Master.Sheets("DATA").Range("G2:G" & LastRowMasterData).ClearContents
        Master.Sheets("DATA").AutoFilter.ShowAllData
        LastRowMasterInfo = Master.Sheets("No. products").Range("A" & Rows.Count).End(xlUp).Row
        
        FutureMonthName = Format(DateSerial(Year(Date), Month(Date) + 1, 1), "mmmm")
        
        If FutureMonthName = "January" Then
            ACTIVERANGE = "D"
        End If
        If FutureMonthName = "February" Then
            ACTIVERANGE = "F"
        End If
        If FutureMonthName = "March" Then
            ACTIVERANGE = "H"
        End If
        If FutureMonthName = "April" Then
            ACTIVERANGE = "J"
        End If
        If FutureMonthName = "May" Then
            ACTIVERANGE = "L"
        End If
        If FutureMonthName = "June" Then
            ACTIVERANGE = "N"
        End If
        If FutureMonthName = "July" Then
            ACTIVERANGE = "P"
        End If
        If FutureMonthName = "August" Then
            ACTIVERANGE = "R"
        End If
        If FutureMonthName = "September" Then
            ACTIVERANGE = "T"
        End If
        If FutureMonthName = "October" Then
            ACTIVERANGE = "V"
        End If
        If FutureMonthName = "November" Then
            ACTIVERANGE = "X"
        End If
        If FutureMonthName = "December" Then
            ACTIVERANGE = "Z"
        End If
        
        
    
        LastRowMasterEstimators = Master.Sheets("DATA").Range("F" & Rows.Count).End(xlUp).Row
        EstimatorsStart = 2
        EstimatorsFinish = LastRowMasterEstimators
        For I = EstimatorsStart To EstimatorsFinish
            InputAccountIdentifier = Master.Sheets("DATA").Range("D" & EstimatorsStart).Value
            SearchEstimator = Master.Sheets("DATA").Range("F" & EstimatorsStart).Value
            Set PriceFile = Workbooks.Open(SearchEstimator, ReadOnly:=True)
            LastRowPriceFile = PriceFile.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
            NOPRODUCTS = Master.Sheets("No. Products").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            NOEOL = Master.Sheets("No. EOL").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            NONOUSAGE = Master.Sheets("No. No-Usage").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            NOINTERNAL = Master.Sheets("No. Internal").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            NONOTINLINESPIN = Master.Sheets("No. Not-Inline-Spin").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            NONOTPASS = Master.Sheets("No. Not Pass").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            LOSSVALUE = Master.Sheets("Loss Value").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            BASKETMARGIN = Master.Sheets("Basket Margin").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            BASKETPROFIT = Master.Sheets("Basket Profit").Range("C:C").Find(InputAccountIdentifier, lookat:=xlPart).Row
            Master.Sheets("No. Products").Range(ACTIVERANGE & NOPRODUCTS).Formula = "=COUNTA('[" & PriceFile.Name & "]Price File'!$A$12:$A$" & LastRowPriceFile & ")"
            Master.Sheets("No. EOL").Range(ACTIVERANGE & NOEOL).Formula = "=COUNTIFS('[" & PriceFile.Name & "]Price File'!$BG:$BG,""EOL"")"
            Master.Sheets("No. No-Usage").Range(ACTIVERANGE & NONOUSAGE).Formula = "=COUNTIFS('[" & PriceFile.Name & "]Price File'!$AQ:$AQ,""No Usage"")"
            Master.Sheets("No. Internal").Range(ACTIVERANGE & NOINTERNAL).Formula = "=COUNTIFS('[" & PriceFile.Name & "]Price File'!$BW:$BW,""*intern*"")"
            Master.Sheets("No. Not-Inline-Spin").Range(ACTIVERANGE & NONOTINLINESPIN).Formula = "=COUNTIFS('[" & PriceFile.Name & "]Price File'!$BW:$BW,""*Not Inline Spin*"")"
            Master.Sheets("No. Not Pass").Range(ACTIVERANGE & NONOTPASS).Formula = "=COUNTIFS('[" & PriceFile.Name & "]Price File'!$AV:$AV,""X"")"
            Master.Sheets("Basket Margin").Range(ACTIVERANGE & BASKETMARGIN).Value = "='[" & PriceFile.Name & "]Price File'!$AU$4"
            Master.Sheets("Basket Profit").Range(ACTIVERANGE & BASKETPROFIT).Value = "='[" & PriceFile.Name & "]Price File'!$AU$5"
            PriceFile.Sheets("Price File").Range("A11:BX" & LastRowMasterData).AutoFilter field:=47, Criteria1:="<0"
            PriceFile.Sheets("Price File").Range("A11:BX" & LastRowMasterData).AutoFilter field:=43, Criteria1:=">0"
            If PriceFile.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                PriceFile.Sheets("Price File").Range("BY12:BY" & LastRowPriceFile).SpecialCells(xlCellTypeVisible).Formula = "=(@AT:AT*@AQ:AQ)-(@BQ:BQ*@AQ:AQ)"
                Master.Sheets("Loss Value").Range(ACTIVERANGE & LOSSVALUE).Value = "=SUM('[" & PriceFile.Name & "]Price File'!$BY:$BY)"
            Else
                Master.Sheets("Loss Value").Range(ACTIVERANGE & LOSSVALUE).Value = 0
            End If
            Master.Sheets("No. Products").Range(ACTIVERANGE & NOPRODUCTS).Copy
            Master.Sheets("No. Products").Range(ACTIVERANGE & NOPRODUCTS).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("No. EOL").Range(ACTIVERANGE & NOEOL).Copy
            Master.Sheets("No. EOL").Range(ACTIVERANGE & NOEOL).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("No. No-Usage").Range(ACTIVERANGE & NONOUSAGE).Copy
            Master.Sheets("No. No-Usage").Range(ACTIVERANGE & NONOUSAGE).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("No. Internal").Range(ACTIVERANGE & NOINTERNAL).Copy
            Master.Sheets("No. Internal").Range(ACTIVERANGE & NOINTERNAL).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("No. Not-Inline-Spin").Range(ACTIVERANGE & NONOTINLINESPIN).Copy
            Master.Sheets("No. Not-Inline-Spin").Range(ACTIVERANGE & NONOTINLINESPIN).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("No. Not Pass").Range(ACTIVERANGE & NONOTPASS).Copy
            Master.Sheets("No. Not Pass").Range(ACTIVERANGE & NONOTPASS).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Loss Value").Range(ACTIVERANGE & LOSSVALUE).Copy
            Master.Sheets("Loss Value").Range(ACTIVERANGE & LOSSVALUE).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Basket Margin").Range(ACTIVERANGE & BASKETMARGIN).Copy
            Master.Sheets("Basket Margin").Range(ACTIVERANGE & BASKETMARGIN).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Basket Profit").Range(ACTIVERANGE & BASKETPROFIT).Copy
            Master.Sheets("Basket Profit").Range(ACTIVERANGE & BASKETPROFIT).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            PriceFile.Close False
            InputAccountIdentifier = 0
            SearchEstimator = 0
            NOPRODUCTS = 0
            NOEOL = 0
            NONOUSAGE = 0
            NOINTERNAL = 0
            NONOTINLINESPIN = 0
            NONOTPASS = 0
            LOSSVALUE = 0
            BASKETMARGIN = 0
            BASKETPROFIT = 0
            
            EstimatorsStart = EstimatorsStart + 1
        Next I
    End If
    
    Master.Sheets("No. Products").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("No. EOL").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("No. No-Usage").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("No. Internal").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("No. Not-Inline-Spin").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("No. Not Pass").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("Loss Value").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("Basket Margin").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    Master.Sheets("Basket Profit").Range("E3:E" & LastRowMasterInfo).Formula = "-"
    
    Master.Sheets("No. Products").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("G3:G" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("I3:I" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("K3:K" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("M3:M" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("O3:O" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("Q3:Q" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("S3:S" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("U3:U" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("W3:W" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("Y3:Y" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    Master.Sheets("No. Products").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. EOL").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. No-Usage").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Internal").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not-Inline-Spin").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("No. Not Pass").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-3],0)"
    Master.Sheets("Loss Value").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Margin").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    Master.Sheets("Basket Profit").Range("AA3:AA" & LastRowMasterInfo).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-3]-1,0)"
    
    
	Master.Sheets("No. Products").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("No. EOL").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("No. No-Usage").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("No. Not-Inline-Spin").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("No. Not Pass").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("Loss Value").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("Basket Margin").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("Basket Profit").Range("A2:AA" & LastRowMasterInfo).AutoFilter
    Master.Sheets("No. Internal").Range("A2:AA" & LastRowMasterInfo).AutoFilter
	
    Master.Sheets("No. Products").Columns.AutoFit
    Master.Sheets("No. Products").Rows.AutoFit
    Master.Sheets("No. EOL").Columns.AutoFit
    Master.Sheets("No. EOL").Rows.AutoFit
    Master.Sheets("No. No-Usage").Columns.AutoFit
    Master.Sheets("No. No-Usage").Rows.AutoFit
    Master.Sheets("No. Internal").Columns.AutoFit
    Master.Sheets("No. Internal").Rows.AutoFit
    Master.Sheets("No. Not-Inline-Spin").Columns.AutoFit
    Master.Sheets("No. Not-Inline-Spin").Rows.AutoFit
    Master.Sheets("No. Not Pass").Columns.AutoFit
    Master.Sheets("No. Not Pass").Rows.AutoFit
    Master.Sheets("Loss Value").Columns.AutoFit
    Master.Sheets("Loss Value").Rows.AutoFit
    Master.Sheets("Basket Margin").Columns.AutoFit
    Master.Sheets("Basket Margin").Rows.AutoFit
    Master.Sheets("Basket Profit").Columns.AutoFit
    Master.Sheets("Basket Profit").Rows.AutoFit
    
        If FutureMonthName = "January" Then
            Master.Sheets("No. Products").Columns("F:Y").Hidden = True
            Master.Sheets("No. EOL").Columns("F:Y").Hidden = True
            Master.Sheets("No. No-Usage").Columns("F:Y").Hidden = True
            Master.Sheets("No. Internal").Columns("F:Y").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("F:Y").Hidden = True
            Master.Sheets("No. Not Pass").Columns("F:Y").Hidden = True
            Master.Sheets("Loss Value").Columns("F:Y").Hidden = True
            Master.Sheets("Basket Margin").Columns("F:Y").Hidden = True
            Master.Sheets("Basket Profit").Columns("F:Y").Hidden = True
        Else
    End If
        If FutureMonthName = "February" Then
            Master.Sheets("No. Products").Columns("H:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("H:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("H:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("H:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("H:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("H:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("H:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("H:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("H:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "March" Then
            Master.Sheets("No. Products").Columns("D:E").Hidden = True
            Master.Sheets("No. EOL").Columns("D:E").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:E").Hidden = True
            Master.Sheets("No. Internal").Columns("D:E").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:E").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:E").Hidden = True
            Master.Sheets("Loss Value").Columns("D:E").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:E").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:E").Hidden = True
            
            Master.Sheets("No. Products").Columns("J:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("J:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("J:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("J:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("J:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("J:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("J:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("J:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("J:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "April" Then
            Master.Sheets("No. Products").Columns("D:G").Hidden = True
            Master.Sheets("No. EOL").Columns("D:G").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:G").Hidden = True
            Master.Sheets("No. Internal").Columns("D:G").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:G").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:G").Hidden = True
            Master.Sheets("Loss Value").Columns("D:G").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:G").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:G").Hidden = True
            
            Master.Sheets("No. Products").Columns("L:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("L:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("L:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("L:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("L:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("L:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("L:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("L:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("L:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "May" Then
            Master.Sheets("No. Products").Columns("D:I").Hidden = True
            Master.Sheets("No. EOL").Columns("D:I").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:I").Hidden = True
            Master.Sheets("No. Internal").Columns("D:I").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:I").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:I").Hidden = True
            Master.Sheets("Loss Value").Columns("D:I").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:I").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:I").Hidden = True
            
            Master.Sheets("No. Products").Columns("N:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("N:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("N:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("N:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("N:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("N:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("N:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("N:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("N:AA").Hidden = True
        
        Else
    End If
        If FutureMonthName = "June" Then
            Master.Sheets("No. Products").Columns("D:K").Hidden = True
            Master.Sheets("No. EOL").Columns("D:K").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:K").Hidden = True
            Master.Sheets("No. Internal").Columns("D:K").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:K").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:K").Hidden = True
            Master.Sheets("Loss Value").Columns("D:K").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:K").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:K").Hidden = True
            
            Master.Sheets("No. Products").Columns("P:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("P:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("P:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("P:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("P:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("P:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("P:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("P:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("P:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "July" Then
            Master.Sheets("No. Products").Columns("D:M").Hidden = True
            Master.Sheets("No. EOL").Columns("D:M").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:M").Hidden = True
            Master.Sheets("No. Internal").Columns("D:M").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:M").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:M").Hidden = True
            Master.Sheets("Loss Value").Columns("D:M").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:M").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:M").Hidden = True
            
            Master.Sheets("No. Products").Columns("R:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("R:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("R:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("R:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("R:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("R:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("R:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("R:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("R:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "August" Then
            Master.Sheets("No. Products").Columns("D:O").Hidden = True
            Master.Sheets("No. EOL").Columns("D:O").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:O").Hidden = True
            Master.Sheets("No. Internal").Columns("D:O").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:O").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:O").Hidden = True
            Master.Sheets("Loss Value").Columns("D:O").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:O").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:O").Hidden = True
            
            Master.Sheets("No. Products").Columns("T:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("T:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("T:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("T:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("T:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("T:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("T:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("T:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("T:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "September" Then
            Master.Sheets("No. Products").Columns("D:Q").Hidden = True
            Master.Sheets("No. EOL").Columns("D:Q").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:Q").Hidden = True
            Master.Sheets("No. Internal").Columns("D:Q").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:Q").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:Q").Hidden = True
            Master.Sheets("Loss Value").Columns("D:Q").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:Q").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:Q").Hidden = True
            
            Master.Sheets("No. Products").Columns("V:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("V:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("V:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("V:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("V:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("V:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("V:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("V:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("V:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "October" Then
            Master.Sheets("No. Products").Columns("D:S").Hidden = True
            Master.Sheets("No. EOL").Columns("D:S").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:S").Hidden = True
            Master.Sheets("No. Internal").Columns("D:S").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:S").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:S").Hidden = True
            Master.Sheets("Loss Value").Columns("D:S").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:S").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:S").Hidden = True
            
            Master.Sheets("No. Products").Columns("X:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("X:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("X:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("X:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("X:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("X:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("X:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("X:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("X:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "November" Then
            Master.Sheets("No. Products").Columns("D:U").Hidden = True
            Master.Sheets("No. EOL").Columns("D:U").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:U").Hidden = True
            Master.Sheets("No. Internal").Columns("D:U").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:U").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:U").Hidden = True
            Master.Sheets("Loss Value").Columns("D:U").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:U").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:U").Hidden = True
            
            Master.Sheets("No. Products").Columns("Z:AA").Hidden = True
            Master.Sheets("No. EOL").Columns("Z:AA").Hidden = True
            Master.Sheets("No. No-Usage").Columns("Z:AA").Hidden = True
            Master.Sheets("No. Internal").Columns("Z:AA").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("Z:AA").Hidden = True
            Master.Sheets("No. Not Pass").Columns("Z:AA").Hidden = True
            Master.Sheets("Loss Value").Columns("Z:AA").Hidden = True
            Master.Sheets("Basket Margin").Columns("Z:AA").Hidden = True
            Master.Sheets("Basket Profit").Columns("Z:AA").Hidden = True
        Else
    End If
        If FutureMonthName = "December" Then
            Master.Sheets("No. Products").Columns("D:W").Hidden = True
            Master.Sheets("No. EOL").Columns("D:W").Hidden = True
            Master.Sheets("No. No-Usage").Columns("D:W").Hidden = True
            Master.Sheets("No. Internal").Columns("D:W").Hidden = True
            Master.Sheets("No. Not-Inline-Spin").Columns("D:W").Hidden = True
            Master.Sheets("No. Not Pass").Columns("D:W").Hidden = True
            Master.Sheets("Loss Value").Columns("D:W").Hidden = True
            Master.Sheets("Basket Margin").Columns("D:W").Hidden = True
            Master.Sheets("Basket Profit").Columns("D:W").Hidden = True
        Else
    End If
    Master.Save
End Sub