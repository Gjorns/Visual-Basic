Sub Margin_Sheet_v6()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim MarginSheet As Workbook
    Dim Estimator As String
    Dim accountid As String
    Dim quoteid As String
    Dim savepath As String
    Set MarginSheet = ActiveWorkbook
    
    Estimator = "Andrei Polin"
    
    If MarginSheet.Sheets(1).Range("A1") = "Branch" Then
        MarginSheet.Activate
        For Each Sheet In MarginSheet.Worksheets
             If Sheet.Name = "Finance Margins" Then
                  Sheet.Delete
             End If
        Next Sheet
        MarginSheet.Sheets(1).Name = "PricePoint"
        MarginSheet.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "Finance Margins"
        LastRowPricePoint = MarginSheet.Sheets("PricePoint").Range("A" & Rows.Count).End(xlUp).Row
        
        MarginSheet.Sheets("Finance Margins").Range("A9").Value = "Finance Margin Sheet - Generated via PricePoint"
        MarginSheet.Sheets("Finance Margins").Range("G9").Value = "Finance Margins"
        MarginSheet.Sheets("Finance Margins").Range("N9").Value = "Basket Totals"
        MarginSheet.Sheets("Finance Margins").Range("Q9").Value = "Cost Plus Calculator"
        MarginSheet.Sheets("Finance Margins").Range("T9").Value = "Discount Off Calculator"
        MarginSheet.Sheets("Finance Margins").Range("W9").Value = "Product Details"
        MarginSheet.Sheets("Finance Margins").Range("Z9").Value = "Pricing Details"
        MarginSheet.Sheets("Finance Margins").Range("A1").Value = "Account ID:"
        MarginSheet.Sheets("Finance Margins").Range("A2").Value = "Account Name:"
        MarginSheet.Sheets("Finance Margins").Range("A3").Value = "Sheet Generator:"
        MarginSheet.Sheets("Finance Margins").Range("A4").Value = "Quote ID:"
        MarginSheet.Sheets("Finance Margins").Range("A5").Value = "Date Generated:"
        MarginSheet.Sheets("Finance Margins").Range("F1").Value = "Total Sell:"
        MarginSheet.Sheets("Finance Margins").Range("F2").Value = "Total Cost:"
        MarginSheet.Sheets("Finance Margins").Range("F3").Value = "Basket Margin:"
        MarginSheet.Sheets("Finance Margins").Range("F4").Value = "Total Profit:"
        MarginSheet.Sheets("Finance Margins").Range("A10").Value = "Supplier Part No."
        MarginSheet.Sheets("Finance Margins").Range("B10").Value = "Wolseley Code"
        MarginSheet.Sheets("Finance Margins").Range("C10").Value = "Product Description"
        MarginSheet.Sheets("Finance Margins").Range("D10").Value = "Terms Price (£)"
        MarginSheet.Sheets("Finance Margins").Range("E10").Value = "Rebated Price (£)"
        MarginSheet.Sheets("Finance Margins").Range("F10").Value = "Quantity"
        MarginSheet.Sheets("Finance Margins").Range("G10").Value = "Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("H10").Value = "Rebated Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("I10").Value = "Cust Rebate (%)"
        MarginSheet.Sheets("Finance Margins").Range("J10").Value = "Support (£)"
        MarginSheet.Sheets("Finance Margins").Range("K10").Value = "WUK Support Ref."
        MarginSheet.Sheets("Finance Margins").Range("L10").Value = "Nett Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("M10").Value = "Nett Cost NRA (£)"
        MarginSheet.Sheets("Finance Margins").Range("N10").Value = "Total Sell (£)"
        MarginSheet.Sheets("Finance Margins").Range("O10").Value = "Total Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("P10").Value = "Total Profit (£)"
        MarginSheet.Sheets("Finance Margins").Range("Q10").Value = "Set Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("R10").Value = "New Sell (£)"
        MarginSheet.Sheets("Finance Margins").Range("S10").Value = "New Discount (%)"
        MarginSheet.Sheets("Finance Margins").Range("T10").Value = "Set Discount (%)"
        MarginSheet.Sheets("Finance Margins").Range("U10").Value = "New Sell (£)"
        MarginSheet.Sheets("Finance Margins").Range("V10").Value = "New Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("W10").Value = "Wolseley SPG"
        MarginSheet.Sheets("Finance Margins").Range("X10").Value = "Supplier Name"
        MarginSheet.Sheets("Finance Margins").Range("Y10").Value = "Supplier HoS"
        MarginSheet.Sheets("Finance Margins").Range("Z10").Value = "Current Invoice (£)"
        MarginSheet.Sheets("Finance Margins").Range("AA10").Value = "Future Invoice (£)"
        MarginSheet.Sheets("Finance Margins").Range("AB10").Value = "Invoice Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AC10").Value = "Current Trade (£)"
        MarginSheet.Sheets("Finance Margins").Range("AD10").Value = "Future Trade (£)"
        MarginSheet.Sheets("Finance Margins").Range("AE10").Value = "Trade Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AF10").Value = "Current Branch (£)"
        MarginSheet.Sheets("Finance Margins").Range("AG10").Value = "Future Branch (£)"
        MarginSheet.Sheets("Finance Margins").Range("AH10").Value = "Branch Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AI10").Value = "Current SLP (£)"
        MarginSheet.Sheets("Finance Margins").Range("AJ10").Value = "Future SLP (£)"
        MarginSheet.Sheets("Finance Margins").Range("AK10").Value = "SLP Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AL10").Value = "Current Real Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("AM10").Value = "Future Real Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("AN10").Value = "Real Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AO10").Value = "Future Date"
        MarginSheet.Sheets("Finance Margins").Range("AP10").Value = "Rebate Impacted"
        MarginSheet.Sheets("Finance Margins").Range("AQ10").Value = "Product Lifecycle"
        MarginSheet.Sheets("Finance Margins").Range("AR10").Value = "Product Narrative"
        MarginSheet.Sheets("Finance Margins").Range("AS10").Value = "Last Sale Date"
        MarginSheet.Sheets("Finance Margins").Range("AT10").Value = "Last Sale Price (£)"
        MarginSheet.Sheets("Finance Margins").Range("H1").Value = "No. of Increases:"
        MarginSheet.Sheets("Finance Margins").Range("H2").Value = "No. of OBS/EOL:"
        MarginSheet.Sheets("Finance Margins").Range("H3").Value = "No. of Loss Making:"
        MarginSheet.Sheets("Finance Margins").Range("H4").Value = "No. of Supports:"
        
        'Get WUK Codes in for Lookup
        MarginSheet.Sheets("PricePoint").Range("H15:H" & LastRowPricePoint).Copy
        MarginSheet.Sheets("Finance Margins").Range("B11").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LastRowFinanceMargins = MarginSheet.Sheets("Finance Margins").Range("B" & Rows.Count).End(xlUp).Row
        
        'Margin Sheet Lookups START
        MarginSheet.Sheets("Finance Margins").Range("A11:A" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!E:E,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("C11:C" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!I:I,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("D11:D" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!AB:AB,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("I11:I" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!AQ:AQ,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("J11:J" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!AN:AN,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("K11:K" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!AJ:AJ,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,"""",INDEX(PricePoint!AJ:AJ,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("W11:W" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!G:G,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("X11:X" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!D:D,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("Z11:Z" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!Q:Q,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AC11:AC" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!O:O,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AF11:AF" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!P:P,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AI11:AI" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!R:R,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AL11:AL" & LastRowFinanceMargins).Formula = "=INDEX(PricePoint!S:S,MATCH('Finance Margins'!B11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AA11:AA" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!W:W,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,@Z:Z,INDEX(PricePoint!W:W,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AD11:AD" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!U:U,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,@AC:AC,INDEX(PricePoint!U:U,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AG11:AG" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!V:V,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,@AF:AF,INDEX(PricePoint!V:V,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AJ11:AJ" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!X:X,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,@AI:AI,INDEX(PricePoint!X:X,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!Y:Y,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,@AL:AL,INDEX(PricePoint!Y:Y,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!T:T,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,""No Increases"",INDEX(PricePoint!T:T,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AP11:AP" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!F:F,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,""No Data"",INDEX(PricePoint!F:F,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AQ11:AQ" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!J:J,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,""No Data"",INDEX(PricePoint!J:J,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AR11:AR" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!M:M,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,""No Data"",INDEX(PricePoint!M:M,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AS11:AS" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!K:K,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,""No Sales"",INDEX(PricePoint!K:K,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AT11:AT" & LastRowFinanceMargins).Formula = "=IF(INDEX(PricePoint!L:L,MATCH('Finance Margins'!B11,PricePoint!H:H,0))=0,""No Sales"",INDEX(PricePoint!L:L,MATCH('Finance Margins'!B11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("B1").Value = MarginSheet.Sheets("PricePoint").Range("B2").Value
        MarginSheet.Sheets("Finance Margins").Range("B2").Value = MarginSheet.Sheets("PricePoint").Range("B3").Value
        MarginSheet.Sheets("Finance Margins").Range("B3").Value = Estimator
        MarginSheet.Sheets("Finance Margins").Range("B4").Value = MarginSheet.Sheets("PricePoint").Range("B4").Value
        MarginSheet.Sheets("Finance Margins").Range("B5").Value = Now()
        
        'Margin Sheet Remove Lookups START
        MarginSheet.Sheets("Finance Margins").Range("A11:AT" & LastRowFinanceMargins).Copy
        MarginSheet.Sheets("Finance Margins").Range("A11:AT" & LastRowFinanceMargins).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        'Margin Sheet Formulas START
        MarginSheet.Sheets("Finance Margins").Range("F11:F" & LastRowFinanceMargins).Value = "1"
        MarginSheet.Sheets("Finance Margins").Range("E11:E" & LastRowFinanceMargins).Formula = "=@D:D*(1-@I:I)"
        MarginSheet.Sheets("Finance Margins").Range("G11:G" & LastRowFinanceMargins).Formula = "=(@D:D-@M:M)/@D:D"
        MarginSheet.Sheets("Finance Margins").Range("H11:H" & LastRowFinanceMargins).Formula = "=(@E:E-@M:M)/@E:E"
        MarginSheet.Sheets("Finance Margins").Range("L11:L" & LastRowFinanceMargins).Formula = "=@AL:AL"
        MarginSheet.Sheets("Finance Margins").Range("M11:M" & LastRowFinanceMargins).Formula = "=IF(@AP:AP=""Y"",@L:L,@L:L-@J:J)"
        MarginSheet.Sheets("Finance Margins").Range("N11:N" & LastRowFinanceMargins).Formula = "=@E:E*@F:F"
        MarginSheet.Sheets("Finance Margins").Range("O11:O" & LastRowFinanceMargins).Formula = "=@M:M*@F:F"
        MarginSheet.Sheets("Finance Margins").Range("P11:P" & LastRowFinanceMargins).Formula = "=@N:N-@O:O"
        MarginSheet.Sheets("Finance Margins").Range("R11:R" & LastRowFinanceMargins).Formula = "=@M:M/(1-@Q:Q)"
        MarginSheet.Sheets("Finance Margins").Range("S11:S" & LastRowFinanceMargins).Formula = "=(@R:R/@AC:AC)-1"
        MarginSheet.Sheets("Finance Margins").Range("U11:U" & LastRowFinanceMargins).Formula = "=@AC:AC*(1-@T:T)"
        MarginSheet.Sheets("Finance Margins").Range("V11:V" & LastRowFinanceMargins).Formula = "=(@U:U-@M:M)/@U:U"
        MarginSheet.Sheets("Finance Margins").Range("AB11:AB" & LastRowFinanceMargins).Formula = "=IFERROR((@AA:AA/@Z:Z)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AE11:AE" & LastRowFinanceMargins).Formula = "=IFERROR((@AD:AD/@AC:AC)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AH11:AH" & LastRowFinanceMargins).Formula = "=IFERROR((@AG:AG/@AF:AF)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AK11:AK" & LastRowFinanceMargins).Formula = "=IFERROR((@AJ:AJ/@AI:AI)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AN11:AN" & LastRowFinanceMargins).Formula = "=IFERROR((@AM:AM/@AL:AL)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("G1").Formula = "=SUM(N11:N" & LastRowFinanceMargins & ")"
        MarginSheet.Sheets("Finance Margins").Range("G2").Formula = "=SUM(O11:O" & LastRowFinanceMargins & ")"
        MarginSheet.Sheets("Finance Margins").Range("G3").Formula = "=(G1-G2)/G1"
        MarginSheet.Sheets("Finance Margins").Range("G4").Formula = "=G1*G3"
        MarginSheet.Sheets("Finance Margins").Range("I1").Formula = "=COUNTIF(AO11:AO" & LastRowFinanceMargins & ",""<>No Increases"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("I2").Formula = "=COUNTIF(AQ11:AQ" & LastRowFinanceMargins & ",""*EOL*"")+COUNTIF(AQ11:AQ" & LastRowFinanceMargins & ",""*OBS*"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("I3").Formula = "=COUNTIF(H11:H" & LastRowFinanceMargins & ",""<0"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("I4").Formula = "=COUNTIF(J11:J" & LastRowFinanceMargins & ","">0"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("C5").Formula = "=TRIM(RIGHT(SUBSTITUTE(B4,""."",REPT("" "",100)),100))"
        
        'Margin Sheet Design START
        MarginSheet.Sheets("Finance Margins").Range("A9:F9,G9:M9,N9:P9,Q9:S9,T9:V9,W9:Y9,Z9:AT9").Merge
        MarginSheet.Sheets("Finance Margins").Range("D11:D" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("E11:E" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("J11:J" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("L11:L" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("M11:M" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("N11:N" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("O11:O" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("P11:P" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("R11:R" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("U11:U" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("Z11:Z" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AA11:AA" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AC11:AC" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AD11:AD" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AF11:AF" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AG11:AG" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AI11:AI" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AJ11:AJ" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AL11:AL" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AT11:AT" & LastRowFinanceMargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G1").NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G2").NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G4").NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G11:G" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("H11:H" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("I11:I" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("Q11:Q" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("S11:S" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("T11:T" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("V11:V" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AB11:AB" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AE11:AE" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AH11:AH" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AK11:AK" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AN11:AN" & LastRowFinanceMargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("G3").NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AS11:AS" & LastRowFinanceMargins).NumberFormat = "dd/mm/yyyy"
        MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & LastRowFinanceMargins).NumberFormat = "dd/mm/yyyy"
        MarginSheet.Sheets("Finance Margins").Range("A9:AT" & LastRowFinanceMargins).HorizontalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("A9:AT" & LastRowFinanceMargins).VerticalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("G1:G4,I1:I4").HorizontalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("G1:G4,I1:I4").VerticalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("A9:AT" & LastRowFinanceMargins).Borders.LineStyle = xlContinuous
        MarginSheet.Sheets("Finance Margins").Range("F1:I4").Borders.LineStyle = xlContinuous
        MarginSheet.Sheets("Finance Margins").Range("A1:B5").Borders.LineStyle = xlContinuous
        MarginSheet.Sheets("Finance Margins").Cells.Font.Name = "Arial"
        MarginSheet.Sheets("Finance Margins").Cells.Font.Size = "10"
        MarginSheet.Sheets("Finance Margins").Range("A1:A5,F1:F4,A9:AT10,H1:H4").Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("G11:G" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("H11:H" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("I11:I" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("S11:S" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("V11:V" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AB11:AB" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AE11:AE" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AH11:AH" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AK11:AK" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AN11:AN" & LastRowFinanceMargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("A10:AT10").Font.Underline = True
        ActiveWindow.DisplayGridlines = False
        MarginSheet.Sheets("Finance Margins").Range("A10:AT10,A1:B5,F1:F4,H1:H4").Interior.Color = RGB(150, 194, 230)
        MarginSheet.Sheets("Finance Margins").Range("A9:F9,G1:G4,I1:I4").Interior.Color = RGB(237, 237, 237)
        MarginSheet.Sheets("Finance Margins").Range("A11:F" & LastRowFinanceMargins).Interior.Color = RGB(237, 237, 237)
        MarginSheet.Sheets("Finance Margins").Range("G9:M9").Interior.Color = RGB(198, 224, 180)
        MarginSheet.Sheets("Finance Margins").Range("G11:M" & LastRowFinanceMargins).Interior.Color = RGB(198, 224, 180)
        MarginSheet.Sheets("Finance Margins").Range("N9:P9").Interior.Color = RGB(146, 208, 80)
        MarginSheet.Sheets("Finance Margins").Range("N11:P" & LastRowFinanceMargins).Interior.Color = RGB(146, 208, 80)
        MarginSheet.Sheets("Finance Margins").Range("Q9:S9").Interior.Color = RGB(255, 204, 204)
        MarginSheet.Sheets("Finance Margins").Range("Q11:S" & LastRowFinanceMargins).Interior.Color = RGB(255, 204, 204)
        MarginSheet.Sheets("Finance Margins").Range("T9:V9").Interior.Color = RGB(248, 203, 173)
        MarginSheet.Sheets("Finance Margins").Range("T11:V" & LastRowFinanceMargins).Interior.Color = RGB(248, 203, 173)
        MarginSheet.Sheets("Finance Margins").Range("W9:AT9").Interior.Color = RGB(242, 242, 242)
        MarginSheet.Sheets("Finance Margins").Range("W11:AT" & LastRowFinanceMargins).Interior.Color = RGB(242, 242, 242)
        MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter
        MarginSheet.Sheets("Finance Margins").Columns("A:AT").AutoFit
        MarginSheet.Sheets("Finance Margins").Range("D10:D" & LastRowFinanceMargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("Q10:Q" & LastRowFinanceMargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("F10:F" & LastRowFinanceMargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("T10:T" & LastRowFinanceMargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("D11:D" & LastRowFinanceMargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("F11:F" & LastRowFinanceMargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("Q11:Q" & LastRowFinanceMargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("T11:T" & LastRowFinanceMargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("B11:B" & LastRowFinanceMargins).TextToColumns Destination:=Range("B11"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        
        'Specialised Logic
        MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=42, Criteria1:="Y"
        MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=11, Criteria1:="<>"
        MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=37, Criteria1:=">0"
        If MarginSheet.Sheets("Finance Margins").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            MarginSheet.Sheets("Finance Margins").Range("AU11:AU" & LastRowFinanceMargins).SpecialCells(xlCellTypeVisible).Formula = "=@AM:AM-@J:J"
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
            MarginSheet.Sheets("Finance Margins").Range("AU11:AU" & LastRowFinanceMargins).Copy
            MarginSheet.Sheets("Finance Margins").Range("AU11:AU" & LastRowFinanceMargins).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=42, Criteria1:="Y"
            MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=11, Criteria1:="<>"
            MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=37, Criteria1:=">0"
            MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & LastRowFinanceMargins).SpecialCells(xlCellTypeVisible).Formula = "=@AU:AU"
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
            MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & LastRowFinanceMargins).Copy
            MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & LastRowFinanceMargins).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MarginSheet.Sheets("Finance Margins").Range("AU11:AU" & LastRowFinanceMargins).ClearContents
        Else
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
        End If
        
        MarginSheet.Sheets("Finance Margins").Range("A10:AT" & LastRowFinanceMargins).AutoFilter Field:=42, Criteria1:="Y"
        If MarginSheet.Sheets("Finance Margins").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            MarginSheet.Sheets("Finance Margins").Range("J11:J" & LastRowFinanceMargins).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 102, 0)
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
        Else
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
        End If
        
        'Deletion of PricePoint and Saving File into specific Folder
        MarginSheet.Sheets("PricePoint").Delete
        
        accountid = MarginSheet.Sheets("Finance Margins").Range("B1").Value
        quoteid = MarginSheet.Sheets("Finance Margins").Range("C5").Value
        MarginSheet.Sheets("Finance Margins").Range("C5").ClearContents
        savepath = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Quotes\" & accountid & " PP" & quoteid & " - " & "Margin Sheet.xlsx"
        MarginSheet.SaveAs savepath
        MarginSheet.Close
    Else
        MsgBox "Raw File does not appear to be PricePoint Download"
    End If
    
End Sub
