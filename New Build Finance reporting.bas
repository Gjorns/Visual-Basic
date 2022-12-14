Sub Step1_RawFilePreparation()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook
    
    Dim RawFile As Workbook
    Dim PriceFile As Workbook
    Dim SPGFile As Workbook
    Dim Removals As Workbook
    Dim Allocations As Workbook
    Dim ManualReview As Workbook
    
    Dim RawFileName As String
    Dim PriceFileName As String
    Dim SPGFileName As String
    Dim RemovalsName As String
    Dim AllocationsName As String
    Dim TodayDate As String
    Dim TodayYear As String
    Dim Manualreviewname As String
    
    RawFileName = Master.Sheets("INFO").Range("B1").Value
    PriceFileName = Master.Sheets("INFO").Range("B2").Value
    SPGFileName = Master.Sheets("INFO").Range("B3").Value
    RemovalsName = Master.Sheets("INFO").Range("B4").Value
    AllocationsName = Master.Sheets("INFO").Range("B5").Value
    TodayDate = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd mmmm yy")
    TodayYear = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "yyyy")
    Manualreviewname = Master.Sheets("INFO").Range("B6").Value

    
    
    Set RawFile = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Sales and Margin Reports\Raw Developer rebate Reports\Developer rebate report - " & RawFileName & ".xlsx")
    Set PriceFile = Workbooks.Open(PriceFileName)
    Set SPGFile = Workbooks.Open(SPGFileName)
    Set Removals = Workbooks.Open(RemovalsName)
    Set Allocations = Workbooks.Open(AllocationsName)
    Set ManualReview = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Sales and Margin Reports\Manual Review\" & Manualreviewname & " Monthly Review.xlsx")
    
    RawFile.Sheets(1).AutoFilterMode = False
    PriceFile.Sheets(1).AutoFilterMode = False
    SPGFile.Sheets(1).AutoFilterMode = False
    Removals.Sheets(1).AutoFilterMode = False
    Allocations.Sheets(1).AutoFilterMode = False
    ManualReview.Sheets(1).AutoFilterMode = False
    
    If IsEmpty(RawFile.Sheets(1).Range("A1").Value) = True Then
        LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row + 4
        RawFile.Sheets(1).Rows(LastRowRawFile - 3 & ":" & LastRowRawFile).Delete
        RawFile.Sheets(1).Rows("1:3").Delete
        LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Else
        LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    End If
    
    RawFile.Sheets(1).Range("A1").Value = "Customer Account No"
    RawFile.Sheets(1).Range("B1").Value = "Customer Name"
    RawFile.Sheets(1).Range("C1").Value = "Region"
    RawFile.Sheets(1).Range("D1").Value = "Brand"
    RawFile.Sheets(1).Range("E1").Value = "Branch Code"
    RawFile.Sheets(1).Range("F1").Value = "Branch Name"
    RawFile.Sheets(1).Range("G1").Value = "Delivery Address"
    RawFile.Sheets(1).Range("H1").Value = "Customer Contract"
    RawFile.Sheets(1).Range("I1").Value = "Contract Claims Reference"
    RawFile.Sheets(1).Range("J1").Value = "Supplier Name"
    RawFile.Sheets(1).Range("K1").Value = "Product Super Category"
    RawFile.Sheets(1).Range("L1").Value = "Product Category"
    RawFile.Sheets(1).Range("M1").Value = "Product Sub Category"
    RawFile.Sheets(1).Range("N1").Value = "LLSPG Code"
    RawFile.Sheets(1).Range("O1").Value = "LLSPG Desc"
    RawFile.Sheets(1).Range("P1").Value = "Product Code"
    RawFile.Sheets(1).Range("Q1").Value = "Description"
    RawFile.Sheets(1).Range("R1").Value = "Invoice / Credit Note Number"
    RawFile.Sheets(1).Range("S1").Value = "Sales Order / AFC Number"
    RawFile.Sheets(1).Range("T1").Value = "Source Transaction SK"
    RawFile.Sheets(1).Range("U1").Value = "Invoice Date"
    RawFile.Sheets(1).Range("V1").Value = "Sales Quantity"
    RawFile.Sheets(1).Range("W1").Value = "Sales Value"
    RawFile.Sheets(1).Range("X1").Value = "Financial Margin"
    RawFile.Sheets(1).Range("Y1").Value = "Fin Margin %"
    RawFile.Sheets(1).Range("Z1").Value = "Contract Claims Value"
    RawFile.Sheets(1).Range("AA1").Value = "Price File "
    RawFile.Sheets(1).Range("AB1").Value = "Compliance/Non Compliance"
    RawFile.Sheets(1).Range("AC1").Value = "Customer Rebate Value"
    RawFile.Sheets(1).Range("AD1").Value = "Net Revenue"
    RawFile.Sheets(1).Range("AE1").Value = "Net Financial Margin"
    RawFile.Sheets(1).Range("AF1").Value = "Net Fin Margin %"
    RawFile.Sheets(1).Range("AG1").Value = "Exception Flag"
    RawFile.Sheets(1).Range("AH1").Value = "Exception Reason"
    RawFile.Sheets(1).Range("AI1").Value = "Price Derivation Group"
    RawFile.Sheets(1).Range("AJ1").Value = "Price Derivation Code"
    RawFile.Sheets(1).Range("AK1").Value = "Price Derivation Desc"
    RawFile.Sheets(1).Range("AL1").Value = "Date"
    RawFile.Sheets(1).Range("AM1").Value = "Year"
    RawFile.Sheets(1).Range("AN1").Value = "Regional Office"
    RawFile.Sheets(1).Range("AO1").Value = "Site Name"
    RawFile.Sheets(1).Range("AP1").Value = "Data Field 3"
    RawFile.Sheets(1).Range("AQ1").Value = "Data Field 4"
    RawFile.Sheets(1).Range("AR1").Value = "Data Field 5"
    RawFile.Sheets(1).Range("AS1").Value = "Data Field 6"

    RawFile.Sheets(1).Range("A1:AS1").Font.FontStyle = "Arial"
    RawFile.Sheets(1).Range("A1:AS1").Font.Bold = True
    RawFile.Sheets(1).Range("A1:AS1").Font.Size = "10"
    RawFile.Sheets(1).Range("A1:AS1").Interior.Color = RGB(198, 224, 180)
    RawFile.Sheets(1).Range("A1:AS1").BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=vbBlack
    RawFile.Sheets(1).Range("AK1:AK" & LastRowRawFile).Copy
    RawFile.Sheets(1).Range("AL1:AS" & LastRowRawFile).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter
    RawFile.Sheets(1).Cells.HorizontalAlignment = xlCenter
    RawFile.Sheets(1).Cells.VerticalAlignment = xlCenter
    RawFile.Sheets(1).Cells.EntireRow.AutoFit
    RawFile.Sheets(1).Cells.EntireColumn.AutoFit
    
    RawFile.Sheets(1).Columns("A:O").NumberFormat = "@"
    RawFile.Sheets(1).Columns("P:P").TextToColumns Destination:=Range("P1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    RawFile.Sheets(1).Columns("Q:Q").NumberFormat = "@"
    RawFile.Sheets(1).Columns("R:R").TextToColumns Destination:=Range("R1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    RawFile.Sheets(1).Columns("S:S").NumberFormat = "@"
    RawFile.Sheets(1).Columns("T:T").TextToColumns Destination:=Range("T1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    RawFile.Sheets(1).Columns("U:U").NumberFormat = "dd/mm/yyyy"
    RawFile.Sheets(1).Columns("V:V").NumberFormat = "General"
    RawFile.Sheets(1).Columns("W:X").NumberFormat = "$#,##0.00"
    RawFile.Sheets(1).Columns("Y:Y").NumberFormat = "0.00%"
    RawFile.Sheets(1).Columns("Z:Z").NumberFormat = "$#,##0.00"
    RawFile.Sheets(1).Columns("AA:AA").NumberFormat = "General"
    RawFile.Sheets(1).Columns("AB:AB").NumberFormat = "@"
    RawFile.Sheets(1).Columns("AC:AE").NumberFormat = "$#,##0.00"
    RawFile.Sheets(1).Columns("AF:AF").NumberFormat = "0.00%"
    RawFile.Sheets(1).Columns("AG:AK").NumberFormat = "@"
    RawFile.Sheets(1).Columns("AL:AL").NumberFormat = "dd mmmm yyyy"
    RawFile.Sheets(1).Columns("AM:AM").NumberFormat = "@"
    RawFile.Sheets(1).Columns("AN:AS").NumberFormat = "General"
    
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=23, Criteria1:="<=0"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            RawFile.Sheets(1).AutoFilter.ShowAllData
            LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("AA2:AA" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@G:G,'[" & Removals.Name & "]Sheet1'!$A:$A,1,FALSE)"
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=27, Criteria1:="<>#N/A"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            RawFile.Sheets(1).AutoFilter.ShowAllData
            LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
    
    RawFile.Sheets(1).Range("AA2:AA" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@P:P,'[" & PriceFile.Name & "]Price File'!$A:$A,1,FALSE)"
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=27, Criteria1:="#N/A"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("AA2:AA" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@N:N,'[" & SPGFile.Name & "]Sheet1'!$A:$A,1,FALSE)"
            RawFile.Sheets(1).AutoFilter.ShowAllData
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=27, Criteria1:="#N/A"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("AA2:AA" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Value = "Non-Price File"
            RawFile.Sheets(1).AutoFilter.ShowAllData
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=27, Criteria1:="<>Non-Price File"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("AA2:AA" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Value = "Price File"
            RawFile.Sheets(1).AutoFilter.ShowAllData
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("AL2:AL" & LastRowRawFile).Value = TodayDate
    RawFile.Sheets(1).Range("AM2:AM" & LastRowRawFile).Value = TodayYear
    
    RawFile.Sheets(1).Range("AN2:AN" & LastRowRawFile).Formula = "=IFERROR(VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$G,2,FALSE),""New Address"")"
    RawFile.Sheets(1).Range("AO2:AO" & LastRowRawFile).Formula = "=IFERROR(VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$G,3,FALSE),""New Address"")"
    RawFile.Sheets(1).Range("AP2:AP" & LastRowRawFile).Formula = "=IFERROR(VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$G,4,FALSE),""New Address"")"
    RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).Copy
    RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    LastRowManualreview = ManualReview.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    ManualReview.Sheets(1).Range("A1:AS" & LastRowManualreview).AutoFilter field:=1, Criteria1:="="
        If ManualReview.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            ManualReview.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            ManualReview.Sheets(1).AutoFilter.ShowAllData
        Else
            ManualReview.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=7, Criteria1:="="
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("AN2:AN" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
            RawFile.Sheets(1).Range("AO2:AN" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
            RawFile.Sheets(1).Range("AP2:AN" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
            RawFile.Sheets(1).AutoFilter.ShowAllData
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
    
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=40, Criteria1:="New Address"
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=7, Criteria1:="<>"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("A2:AS" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Copy
            ManualReview.Sheets(1).Range("A2").PasteSpecial Paste:=xlPasteFormats
            ManualReview.Sheets(1).Range("A2").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowManualreview = ManualReview.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
            ManualReview.Sheets(1).Range("A1:AS" & LastRowManualreview).RemoveDuplicates Columns:=7, Header:=xlYes
            RawFile.Sheets(1).AutoFilter.ShowAllData
            ManualReview.Sheets(1).Cells.EntireRow.AutoFit
            ManualReview.Sheets(1).Cells.EntireColumn.AutoFit
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If

    RawFile.Sheets(1).Cells.EntireRow.AutoFit
    RawFile.Sheets(1).Cells.EntireColumn.AutoFit
    
    ManualReview.Close True
    SPGFile.Close False
    PriceFile.Close False
    Removals.Close False
    Allocations.Close False
    RawFile.Close True
End Sub

Sub Step2_AfterManualReview()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook

    Dim Removals As Workbook
    Dim RawFile As Workbook
    Dim Allocations As Workbook
    Dim ManualReview As Workbook
    
    Dim Manualreviewname As String
    Dim RemovalsName As String
    Dim AllocationsName As String
    
    RemovalsName = Master.Sheets("INFO").Range("B4").Value
    RawFileName = Master.Sheets("INFO").Range("B1").Value
    AllocationsName = Master.Sheets("INFO").Range("B5").Value
    Manualreviewname = Master.Sheets("INFO").Range("B6").Value
    
    Set Removals = Workbooks.Open(RemovalsName)
    Set RawFile = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Sales and Margin Reports\Raw Developer rebate Reports\Developer rebate report - " & RawFileName & ".xlsx")
    Set Allocations = Workbooks.Open(AllocationsName)
    Set ManualReview = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Sales and Margin Reports\Manual Review\" & Manualreviewname & " Monthly Review.xlsx")
    
    Removals.Sheets(1).AutoFilterMode = False
    RawFile.Sheets(1).AutoFilterMode = False
    Allocations.Sheets(1).AutoFilterMode = False
    ManualReview.Sheets(1).AutoFilterMode = False
    
    LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    LastRowRemovals = Removals.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    LastRowAllocations = Allocations.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    LastRowManualreview = ManualReview.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
    ManualReview.Sheets(1).Range("A1:AS" & LastRowManualreview).AutoFilter field:=45, Criteria1:="REMOVE"
        If ManualReview.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            ManualReview.Sheets(1).Range("G2:G" & LastRowManualreview).Copy
            Removals.Sheets(1).Range("A" & LastRowRemovals + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowRemovals = Removals.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
            Removals.Sheets(1).Range("A2:A" & LastRowRemovals).RemoveDuplicates Columns:=1, Header:=xlNo
            ManualReview.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            ManualReview.Sheets(1).AutoFilter.ShowAllData
        Else
            ManualReview.Sheets(1).AutoFilter.ShowAllData
        End If
        
    ManualReview.Sheets(1).Range("A1:AS" & LastRowManualreview).AutoFilter field:=45, Criteria1:="="
        If ManualReview.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            ManualReview.Sheets(1).Range("G2:G" & LastRowManualreview).Copy
            Allocations.Sheets(1).Range("A" & LastRowAllocations + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            ManualReview.Sheets(1).Range("AN2:AP" & LastRowManualreview).Copy
            Allocations.Sheets(1).Range("B" & LastRowAllocations + 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowAllocations = Allocations.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
            Allocations.Sheets(1).Range("A2:G" & LastRowAllocations).RemoveDuplicates Columns:=1, Header:=xlYes
            ManualReview.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            ManualReview.Sheets(1).AutoFilter.ShowAllData
        Else
            ManualReview.Sheets(1).AutoFilter.ShowAllData
        End If
    
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=7, Criteria1:="<>"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("AS2:AS" & LastRowRawFile).Formula = "=VLOOKUP(@G:G,[" & Removals.Name & "]Sheet1!$A:$A,1,FALSE)"
            RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=45, Criteria1:="<>#N/A"
                If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                    RawFile.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
                End If
            RawFile.Sheets(1).AutoFilter.ShowAllData
            LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
            RawFile.Sheets(1).Range("AS2:AS" & LastRowRawFile).ClearContents
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=7, Criteria1:="<>"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets(1).Range("AN2:AN" & LastRowRawFile).Formula = "=VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$D,2,FALSE)"
            RawFile.Sheets(1).Range("AO2:AO" & LastRowRawFile).Formula = "=VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$D,3,FALSE)"
            RawFile.Sheets(1).Range("AP2:AP" & LastRowRawFile).Formula = "=VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$D,4,FALSE)"
            RawFile.Sheets(1).AutoFilter.ShowAllData
            RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).Copy
            RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets(1).Range("AS2:AS" & LastRowRawFile).FormulaR1C1 = "=RC[-40] & RC[-44]"
    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=7, Criteria1:="<>"
        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets.Add After:=Sheets(1)
            RawFile.Sheets(1).Range("AS2:AS" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Copy
            RawFile.Sheets(2).Range("A1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Copy
            RawFile.Sheets(2).Range("B1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=7, Criteria1:="="
                If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                    RawFile.Sheets(1).Range("AN2:AN" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@AS:AS,Sheet1!A:D,2,FALSE)"
                    RawFile.Sheets(1).Range("AO2:AO" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@AS:AS,Sheet1!A:D,3,FALSE)"
                    RawFile.Sheets(1).Range("AP2:AP" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@AS:AS,Sheet1!A:D,4,FALSE)"
                    RawFile.Sheets(1).Range("A1:AS" & LastRowRawFile).AutoFilter field:=40, Criteria1:="#N/A"
                        If RawFile.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                            RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
                        End If
                End If
            RawFile.Sheets(1).AutoFilter.ShowAllData
            LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
            RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).Copy
            RawFile.Sheets(1).Range("AN2:AP" & LastRowRawFile).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            RawFile.Sheets(2).Delete
            RawFile.Sheets(1).Range("AS2:AS" & LastRowRawFile).ClearContents
        Else
            RawFile.Sheets(1).AutoFilter.ShowAllData
        End If
    RawFile.Close True
    Allocations.Close True
    ManualReview.Close False
    Removals.Close True
End Sub

Sub Step3_Generator()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim Master As ThisWorkbook
    Set Master = ThisWorkbook
    
    Dim RawFile As Workbook
    Dim AllSalesLarge As Workbook
    Dim Allocations As Workbook
    Dim SmallAllSales As Workbook
    Dim Unknowns As Workbook
    Dim RawFileName As String
    Dim DirYear As String
    Dim DirMonth As String
    Dim PartyPath As String
    Dim PartyName As String
    Dim FullPartyPath As String
    Dim AllSalesSheetName As String
    Dim DeveloperName As String
    Dim StartD As Date
    Dim EndD As Date
    Dim AllocationsName As String
    Dim SmallAllSalesName As String
    Dim UnknownsName As String
    Dim LongMonthName As String
    
    RawFileName = Master.Sheets("INFO").Range("B1").Value
    DirYear = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "yyyy")
    DirMonth = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "MMM YYYY")
    AllSalesSheetName = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "MMM YY")
    PartyPath = "\\wukrls00fp001\RLS_Data\Business Objects\All reps\3RD PARTY REBATE DATA\"
    PartyName = Master.Sheets("Info").Range("B6").Value
    DeveloperName = Master.Sheets("Info").Range("B7").Value
    FullPartyPath = PartyPath & "\" & DirYear & "\" & DirMonth & "\" & PartyName & ".xlsx"
    StartD = CDate(Application.WorksheetFunction.EoMonth(Now, -2))
    EndD = CDate(Application.WorksheetFunction.EoMonth(Now, -1))
    AllocationsName = Master.Sheets("INFO").Range("B5").Value
    SmallAllSalesName = Master.Sheets("INFO").Range("B8").Value
    UnknownsName = Master.Sheets("INFO").Range("B9").Value
    LongMonthName = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mmmm")
    
    Set RawFile = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Sales and Margin Reports\Raw Developer rebate Reports\Developer rebate report - " & RawFileName & ".xlsx")
    RawFile.Sheets(1).AutoFilterMode = False
    LastRowRawFile = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row

    RawFile.Sheets(1).Range("W" & LastRowRawFile + 1).Formula = "=SUM(W2:W" & LastRowRawFile & ")"
    RawFile.Sheets(1).Range("X" & LastRowRawFile + 1).Formula = "=SUM(X2:X" & LastRowRawFile & ")"
    RawFile.Sheets(1).Range("Y" & LastRowRawFile + 1).Formula = "=X" & LastRowRawFile + 1 & "/W" & LastRowRawFile + 1 & ""
    RawFile.Sheets(1).Range("AC" & LastRowRawFile + 1).Formula = "=SUM(AC2:AC" & LastRowRawFile & ")"
    RawFile.Sheets(1).Range("AD" & LastRowRawFile + 1).Formula = "=SUM(AD2:AD" & LastRowRawFile & ")"
    RawFile.Sheets(1).Range("AE" & LastRowRawFile + 1).Formula = "=SUM(AE2:AE" & LastRowRawFile & ")"
    RawFile.Sheets(1).Range("AF" & LastRowRawFile + 1).Formula = "=AE" & LastRowRawFile + 1 & "/AD" & LastRowRawFile + 1 & ""
    RawFile.Sheets(1).Range("A" & LastRowRawFile + 1 & ":AS" & LastRowRawFile + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=vbBlack
    RawFile.Sheets(1).Range("A" & LastRowRawFile + 1 & ":AS" & LastRowRawFile + 1).Font.Bold = True
    RawFile.Sheets(1).Name = "Completed"
    RawFile.Sheets("Completed").Copy After:=Sheets("Completed")
    ActiveSheet.Name = "Price File"
    RawFile.Sheets("Completed").Copy After:=Sheets("Completed")
    ActiveSheet.Name = "Non-Price File"
    
    LastRowRawFilePF = RawFile.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    LastRowRawFileNONPF = RawFile.Sheets("Non-Price File").Range("A" & Rows.Count).End(xlUp).Row
    RawFile.Sheets("Price File").Range("A1:AS" & LastRowRawFilePF).AutoFilter field:=27, Criteria1:="<>Price File"
        If RawFile.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets("Price File").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            RawFile.Sheets("Price File").AutoFilter.ShowAllData
            LastRowRawFilePF = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Else
            RawFile.Sheets("Price File").AutoFilter.ShowAllData
        End If
        
    RawFile.Sheets("Non-Price File").Range("A1:AS" & LastRowRawFileNONPF).AutoFilter field:=27, Criteria1:="<>Non-Price File"
        If RawFile.Sheets("Non-Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            RawFile.Sheets("Non-Price File").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            RawFile.Sheets("Non-Price File").AutoFilter.ShowAllData
            LastRowRawFileNONPF = RawFile.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Else
            RawFile.Sheets("Non-Price File").AutoFilter.ShowAllData
        End If
        
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    RawFile.Sheets("Price File").Activate
        Lastrowpivot = RawFile.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    SrcData = ActiveSheet.Name & "!" & Range("A1:AS" & Lastrowpivot).Address(ReferenceStyle:=xlR1C1)
    Set sht = Sheets.Add
    StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    pvt.PivotFields("Customer Account No").Orientation = xlRowField
    pvt.PivotFields("Customer Name").Orientation = xlRowField
        Dim pF As String
        Dim pf_Name As String
        pF = "Sales Value"
        pf_Name = "Sum of Sales Value"
        Set pvt = ActiveSheet.PivotTables("PivotTable1")
    pvt.AddDataField pvt.PivotFields("Sales Value"), pf_Name, xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sum of Sales Value")
        .NumberFormat = "$#,##0.00"
    End With
    pvt.PivotFields("Sales Value").Subtotals(1) = False
    pvt.RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Sales Value").Subtotals(1) = False
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Customer Account No").Subtotals(1) = False
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Customer Name").Subtotals(1) = False
    ActiveSheet.Name = "Contractor Spend"
    Worksheets("Contractor Spend").Move After:=Worksheets("Price File")
    RawFile.SaveAs PartyPath & DirYear & "\" & DirMonth & "\" & PartyName & " SALES AND MARGIN " & DirMonth & ".xlsx"
    
    Set AllSalesLarge = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Developer Budgets & Salestrack\FY21\All Sales Details by Month 2022.xlsb")
    AllSalesLarge.Sheets(AllSalesSheetName).AutoFilterMode = False
    LastRowRawAllSalesLarge = AllSalesLarge.Sheets(AllSalesSheetName).Range("A" & Rows.Count).End(xlUp).Row
    
    AllSalesLarge.Sheets(AllSalesSheetName).Range("A1:AT" & LastRowRawAllSalesLarge).AutoFilter field:=1, Criteria1:=DeveloperName
        If AllSalesLarge.Sheets(AllSalesSheetName).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            AllSalesLarge.Sheets(AllSalesSheetName).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            AllSalesLarge.Sheets(AllSalesSheetName).AutoFilter.ShowAllData
            LastRowRawAllSalesLarge = AllSalesLarge.Sheets(AllSalesSheetName).Range("A" & Rows.Count).End(xlUp).Row
        Else
            AllSalesLarge.Sheets(AllSalesSheetName).AutoFilter.ShowAllData
        End If
    
    RawFile.Sheets("Completed").Range("A2:AS" & LastRowRawFile).Copy
    AllSalesLarge.Sheets(AllSalesSheetName).Range("B" & LastRowRawAllSalesLarge + 1).PasteSpecial Paste:=xlPasteValues
    AllSalesLarge.Sheets(AllSalesSheetName).Range("B" & LastRowRawAllSalesLarge + 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    LastRowRawAllSalesLargeB = AllSalesLarge.Sheets(AllSalesSheetName).Range("B" & Rows.Count).End(xlUp).Row
    AllSalesLarge.Sheets(AllSalesSheetName).Range("A" & LastRowRawAllSalesLarge + 1 & ":" & "A" & LastRowRawAllSalesLargeB).Value = DeveloperName
    AllSalesLarge.Close True
    
    Set Allocations = Workbooks.Open(AllocationsName)
    Allocations.Sheets(1).AutoFilterMode = False
    LastRowRawAllocations = Allocations.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
    Master.Sheets("DATA").AutoFilterMode = False
    LastRowMaster = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
    Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=38, Criteria1:=">" & CDbl(StartD), Operator:=xlAnd, Criteria2:="<=" & CDbl(EndD)
        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("DATA").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Master.Sheets("DATA").AutoFilter.ShowAllData
            LastRowMaster = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
        Else
            Master.Sheets("DATA").AutoFilter.ShowAllData
        End If
    LastRowRawFilePF = RawFile.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
    RawFile.Sheets("Price File").Range("A2:AS" & LastRowRawFilePF).Copy
    Master.Sheets("DATA").Range("A" & LastRowMaster + 1).PasteSpecial Paste:=xlPasteValues
    Master.Sheets("DATA").Range("A" & LastRowMaster + 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    LastRowMaster = Master.Sheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
    
    Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=7, Criteria1:="<>"
        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("DATA").Range("AN2:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$D,2,FALSE),""UNKNOWN"")"
            Master.Sheets("DATA").Range("AO2:AO" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$D,3,FALSE),""UNKNOWN"")"
            Master.Sheets("DATA").Range("AP2:AP" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=IFERROR(VLOOKUP(@G:G,[" & Allocations.Name & "]Sheet1!$A:$D,4,FALSE),""UNKNOWN"")"
            Master.Sheets("DATA").AutoFilter.ShowAllData
            Master.Sheets("DATA").Range("AN2:AP" & LastRowMaster).Copy
            Master.Sheets("DATA").Range("AN2:AP" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        Else
            Master.Sheets("DATA").AutoFilter.ShowAllData
        End If
    Allocations.Close False
        
    Master.Sheets("DATA").Range("AS2:AS" & LastRowMaster).FormulaR1C1 = "=RC[-40] & RC[-44]"
    Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=7, Criteria1:="<>"
        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets.Add.Name = "TEMP"
            Master.Sheets("DATA").Range("AS2:AS" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
            Master.Sheets("TEMP").Range("A1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("DATA").Range("AN2:AP" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
            Master.Sheets("TEMP").Range("B1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=7, Criteria1:="="
                If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                    Master.Sheets("DATA").Range("AN2:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@AS:AS,TEMP!A:D,2,FALSE)"
                    Master.Sheets("DATA").Range("AO2:AO" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@AS:AS,TEMP!A:D,3,FALSE)"
                    Master.Sheets("DATA").Range("AP2:AP" & LastRowMaster).SpecialCells(xlCellTypeVisible).Formula = "=VLOOKUP(@AS:AS,TEMP!A:D,4,FALSE)"
                    Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=40, Criteria1:="#N/A"
                        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                            Master.Sheets("DATA").Range("AN2:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
                            Master.Sheets("DATA").Range("AO2:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
                            Master.Sheets("DATA").Range("AP2:AN" & LastRowMaster).SpecialCells(xlCellTypeVisible).Value = "Blank Address"
                        End If
                End If
            Master.Sheets("DATA").AutoFilter.ShowAllData
            Master.Sheets("DATA").Range("AN2:AP" & LastRowMaster).Copy
            Master.Sheets("DATA").Range("AN2:AP" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("TEMP").Delete
            Master.Sheets("DATA").Range("AS2:AS" & LastRowMaster).ClearContents
        Else
            Master.Sheets("DATA").AutoFilter.ShowAllData
        End If
        
    Set SmallAllSales = Workbooks.Open(SmallAllSalesName)
    SmallAllSales.Sheets(1).AutoFilterMode = False
    LastRowSmallAllSales = SmallAllSales.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    SmallAllSales.Sheets(1).Range("A1:AS" & LastRowSmallAllSales).AutoFilter field:=38, Criteria1:=">" & CDbl(StartD), Operator:=xlAnd, Criteria2:="<=" & CDbl(EndD)
        If SmallAllSales.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            SmallAllSales.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            SmallAllSales.Sheets(1).AutoFilter.ShowAllData
            LastRowSmallAllSales = SmallAllSales.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Else
            SmallAllSales.Sheets(1).AutoFilter.ShowAllData
        End If
    RawFile.Sheets("Completed").Range("A2:AS" & LastRowRawFile).Copy
    SmallAllSales.Sheets(1).Range("A" & LastRowSmallAllSales + 1).PasteSpecial Paste:=xlPasteValues
    SmallAllSales.Sheets(1).Range("A" & LastRowSmallAllSales + 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    SmallAllSales.Close True
    RawFile.Close True
    
    Set Unknowns = Workbooks.Open(UnknownsName)
    Unknowns.Sheets(1).AutoFilterMode = False
    LastRowUnknowns = Unknowns.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Unknowns.Sheets(1).Range("A1:AS" & LastRowUnknowns).AutoFilter field:=1, Criteria1:="<>"
        If Unknowns.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Unknowns.Sheets(1).AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Unknowns.Sheets(1).AutoFilter.ShowAllData
            LastRowUnknowns = Unknowns.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        Else
            Unknowns.Sheets(1).AutoFilter.ShowAllData
        End If
        
    Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=7, Criteria1:="<>"
    Master.Sheets("DATA").Range("A1:AS" & LastRowMaster).AutoFilter field:=40, Criteria1:="UNKNOWN"
        If Master.Sheets("DATA").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets("DATA").Range("A2:AS" & LastRowMaster).SpecialCells(xlCellTypeVisible).Copy
            Unknowns.Sheets(1).Range("A2").PasteSpecial Paste:=xlPasteValues
            Unknowns.Sheets(1).Range("A2").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            Master.Sheets("DATA").AutoFilter.ShowAllData
        Else
            Master.Sheets("DATA").AutoFilter.ShowAllData
        End If
       
    Master.Sheets("DATA").Columns("A:O").NumberFormat = "@"
    Master.Sheets("DATA").Columns("P:P").TextToColumns Destination:=Range("P1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("DATA").Columns("Q:Q").NumberFormat = "@"
    Master.Sheets("DATA").Columns("R:R").TextToColumns Destination:=Range("R1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("DATA").Columns("S:S").NumberFormat = "@"
    Master.Sheets("DATA").Columns("T:T").TextToColumns Destination:=Range("T1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Master.Sheets("DATA").Columns("U:U").NumberFormat = "dd/mm/yyyy"
    Master.Sheets("DATA").Columns("V:V").NumberFormat = "General"
    Master.Sheets("DATA").Columns("W:X").NumberFormat = "$#,##0.00"
    Master.Sheets("DATA").Columns("Y:Y").NumberFormat = "0.00%"
    Master.Sheets("DATA").Columns("Z:Z").NumberFormat = "$#,##0.00"
    Master.Sheets("DATA").Columns("AA:AA").NumberFormat = "General"
    Master.Sheets("DATA").Columns("AB:AB").NumberFormat = "@"
    Master.Sheets("DATA").Columns("AC:AE").NumberFormat = "$#,##0.00"
    Master.Sheets("DATA").Columns("AF:AF").NumberFormat = "0.00%"
    Master.Sheets("DATA").Columns("AG:AK").NumberFormat = "@"
    Master.Sheets("DATA").Columns("AL:AL").NumberFormat = "dd mmmm yyyy"
    Master.Sheets("DATA").Columns("AM:AM").NumberFormat = "@"
    Master.Sheets("DATA").Columns("AN:AS").NumberFormat = "General"
    Master.RefreshAll
    Unknowns.Close True
    Master.Save
    Master.Sheets("DATA").Range("I2:I" & LastRowMaster).ClearContents
    Master.Sheets("DATA").Range("R2:R" & LastRowMaster).ClearContents
    Master.Sheets("DATA").Range("S2:S" & LastRowMaster).ClearContents
    Master.Sheets("DATA").Range("T2:T" & LastRowMaster).ClearContents
    Master.Sheets("DATA").Range("U2:U" & LastRowMaster).ClearContents
    Master.Sheets("DATA").Range("X2:AK" & LastRowMaster).ClearContents
    Master.RefreshAll
    Master.Sheets("Total Sales").Range("A6:N9").Copy
    Master.Sheets("Total Sales").Range("A6:N9").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Master.Sheets("DATA").Delete
    
    Worksheets(Array("Total Sales", "Site-Contractor-Product", "Category Sales", "Contractor-Cat Sales", "Supplier Split", "Regional Office Split")).Copy
    ActiveWorkbook.SaveAs "\\WUKRLS00FP001\RLS_Data\Departments\NACET\House Builders\Sales and Margin Reports\Customer Version\2022\" & LongMonthName & "\" & PartyName & " Customer Version.xlsx"
    Master.Close False
End Sub